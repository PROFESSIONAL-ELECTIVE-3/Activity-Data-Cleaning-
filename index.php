<?php

$workbookFile = __DIR__ . '/MMORS_water_quality_results_2012-2018_orig.xlsx';
$rivers = ['Marilao', 'Meycauayan', 'Obando'];           // sheet names in the workbook
$combinedLabel = 'All Rivers';
$perPage = 25;

/* Workbook columns G..R (index 6..17) -> output column names */
$paramCols = [
    6  => 'DO_mgL',
    7  => 'pH',
    8  => 'Temperature_C',
    9  => 'BOD_mgL',
    10 => 'TSS_mgL',
    11 => 'Color_TCU',
    12 => 'Fecal_Coliform_MPN',
    13 => 'Total_Coliform_MPN',
    14 => 'Ammonia_mgL',
    15 => 'Nitrates_mgL',
    16 => 'Phosphates_mgL',
    17 => 'Chlorides_mgL',
];
$headers = array_merge(
    ['River', 'Year', 'Month', 'Station_No', 'Station_Name', 'Latitude', 'Longitude', 'Date', 'Time'],
    array_values($paramCols),
    ['Qualifier_Flags']
);

const MONTHS = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];

/* ---------------------- IMPUTATION SETTINGS (NEW) ---------------------------
   Position of each field inside one OUTPUT row (same order as $headers).        */
const COL_YEAR    = 1;    // Year
const COL_MONTH   = 2;    // Month name (JANUARY..DECEMBER)
const COL_STATION = 3;    // Station_No
const COL_DATE    = 7;    // Date  (Y-m-d or '')
const COL_TIME    = 8;    // Time  (HH:MM or '')
const COL_P0      = 9;    // first parameter column (DO_mgL)
const N_PARAMS    = 12;   // DO ... Chlorides  -> columns 9..20
 
/* A gap of up to this many consecutive MISSING months, with a real reading on both
   sides, is filled with the MEAN of the previous and next reading.
   Anything longer is a "large" gap and is filled with the MEDIAN.               */
const MAX_GAP_FOR_MEAN = 3;
 
/* Need at least this many same-month readings (other years) before trusting
   the "same calendar month" median; otherwise use the station's overall median. */
const MIN_SAME_MONTH = 2;
 
/* Fecal/Total Coliform are counts (MPN) -> round to whole numbers (param index 6 and 7). */
const COLIFORM_PARAMS = [6, 7];

/* -------------------------------XLSX reader--------------------------------- */

function colIndex(string $ref): int {
    preg_match('/^([A-Z]+)/', $ref, $m);
    $n = 0;
    foreach (str_split($m[1]) as $ch) $n = $n * 26 + (ord($ch) - 64);
    return $n - 1;
}

/* Returns [sheetName => list of rows], each row = [colIndex => string|float|null]. */
function readWorkbook(string $path): array {
    $zip = new ZipArchive();
    if ($zip->open($path) !== true) throw new RuntimeException('Cannot open workbook.');

    $shared = [];
    $ss = $zip->getFromName('xl/sharedStrings.xml');
    if ($ss !== false) {
        $x = simplexml_load_string($ss);
        foreach ($x->si as $si) {
            if (isset($si->t)) {
                $shared[] = (string)$si->t;
            } else {
                $t = '';
                foreach ($si->r as $r) $t .= (string)$r->t;
                $shared[] = $t;
            }
        }
    }

    $wb   = simplexml_load_string($zip->getFromName('xl/workbook.xml'));
    $rels = simplexml_load_string($zip->getFromName('xl/_rels/workbook.xml.rels'));
    $relMap = [];
    foreach ($rels->Relationship as $rel) {
        $relMap[(string)$rel['Id']] = ltrim((string)$rel['Target'], '/');
    }

    $sheets = [];
    foreach ($wb->sheets->sheet as $sheet) {
        $rid = (string)$sheet->attributes('http://schemas.openxmlformats.org/officeDocument/2006/relationships')['id'];
        if (!isset($relMap[$rid])) continue;
        $target = $relMap[$rid];
        $file = (strpos($target, 'xl/') === 0) ? $target : 'xl/' . $target;
        $xml = $zip->getFromName($file);
        if ($xml === false) continue;

        $sx = simplexml_load_string($xml);
        $rows = [];
        foreach ($sx->sheetData->row as $row) {
            $cells = [];
            foreach ($row->c as $c) {
                $t = (string)$c['t'];
                $val = null;
                if ($t === 's')             $val = $shared[(int)$c->v] ?? null;
                elseif ($t === 'inlineStr')  $val = (string)$c->is->t;
                elseif ($t === 'str')        $val = (string)$c->v;
                elseif ($t === 'b' || $t === 'e') $val = null;
                elseif (isset($c->v) && (string)$c->v !== '') $val = (float)$c->v;
                $cells[colIndex((string)$c['r'])] = $val;
            }
            $rows[] = $cells;
        }
        $sheets[trim((string)$sheet['name'])] = $rows;
    }
    $zip->close();
    return $sheets;
}

/* ---------------------------------cleaning helpers------------------------------- */

/* Returns [float|null, '<'|'>'|null]. '<0.048*' -> half the limit (flagged); '>16000' -> the bound (flagged). */
function cleanMeasure($v): array {
    if ($v === null) return [null, null];
    if (is_float($v) || is_int($v)) return [(float)$v, null];
    $s = preg_replace('/[\s*,]/', '', (string)$v);      // spaces, asterisks, thousands commas ("26. 344" -> 26.344)
    if ($s === '' || $s === '_' || $s === '-') return [null, null];
    if ($s[0] === '<') {
        $lim = substr($s, 1);
        return is_numeric($lim) ? [(float)$lim / 2, '<'] : [null, null];
    }
    if ($s[0] === '>') {
        $lim = substr($s, 1);
        return is_numeric($lim) ? [(float)$lim, '>'] : [null, null];
    }
    return is_numeric($s) ? [(float)$s, null] : [null, null];
}

/* Excel serial number or text like "10/17//2016" -> DateTime|null. */
function parseDate($v): ?DateTime {
    if (is_float($v)) {
        if ($v < 1) return null;
        return (new DateTime('1899-12-30'))->modify('+' . (int)$v . ' days');
    }
    if (is_string($v) && preg_match('#(\d{1,2})/+(\d{1,2})/+(\d{4})#', $v, $m)) {
        if (checkdate((int)$m[1], (int)$m[2], (int)$m[3])) {
            return new DateTime(sprintf('%04d-%02d-%02d', $m[3], $m[1], $m[2]));
        }
    }
    return null;
}

/* Excel time fraction or text like "12:06PM" -> "HH:MM" (24h), or ''. */
function parseTime($v): string {
    if (is_float($v)) {
        if ($v < 0 || $v >= 1) return '';
        $sec = (int)round($v * 86400);
        return sprintf('%02d:%02d', intdiv($sec, 3600) % 24, intdiv($sec % 3600, 60));
    }
    if (is_string($v) && preg_match('/^\s*(\d{1,2}):(\d{2})\s*(AM|PM)\s*$/i', $v, $m)) {
        $h = (int)$m[1] % 12;
        if (strtoupper($m[3]) === 'PM') $h += 12;
        return sprintf('%02d:%s', $h, $m[2]);
    }
    return '';
}

function processSheet(array $rows, string $river): array {
    global $paramCols;
    $stats = ['initial' => 0, 'missing' => 0, 'duplicates' => 0, 'corrected' => 0, 'flagged' => 0];
    $clean = [];
    $seen = [];
    $year = null;
    $month = null;   // 1-12

    foreach ($rows as $cells) {
        $a = $cells[0] ?? null;

        /* Month banner, e.g. "CY 2012 JANUARY" (tolerates typo "CY 20178 DECEMBER") */
        if (is_string($a) && preg_match('/^\s*CY\s*(\d{4})\d*\s+([A-Za-z]+)/', $a, $m)) {
            $idx = array_search(strtoupper($m[2]), MONTHS, true);
            if ($idx !== false) { $year = (int)$m[1]; $month = $idx + 1; }
            continue;
        }

        /* Station rows start with the station number (1..5); everything else is header/guideline/footnote */
        if (!is_float($a) || $a < 1 || floor($a) != $a || $year === null) continue;
        $stats['initial']++;

        /* Date column is authoritative when present (banners have a few typos, e.g. "CY 2016 AUGUST" with 2017 dates) */
        $dt = parseDate($cells[4] ?? null);
        $rowYear = $year;
        $rowMonth = $month;
        if ($dt !== null) {
            $dy = (int)$dt->format('Y');
            $dm = (int)$dt->format('n');
            if ($dy !== $year || $dm !== $month) $stats['corrected']++;
            $rowYear = $dy;
            $rowMonth = $dm;
        }

        
        $values = [];
        $flags = [];
        $hasData = false;
        foreach ($paramCols as $i => $name) {
            [$val, $q] = cleanMeasure($cells[$i] ?? null);
            if ($val !== null) $hasData = true;
            if ($q !== null) $flags[] = $name . ($q === '<' ? '<DL' : '>MAX');
            $values[] = ($val === null) ? null : round($val, 6);
        }

        if (!$hasData) {
            /* CHANGED: a "NO SAMPLING CONDUCTED" row is now COUNTED but KEPT (no `continue`).
               All 12 parameters stay null here and are imputed later in imputeRivers(). */
            $stats['missing']++;
        }
        if ($flags) $stats['flagged']++;

        $lat = isset($cells[2]) && is_float($cells[2]) ? round($cells[2], 6) : null;
        $lon = isset($cells[3]) && is_float($cells[3]) ? round($cells[3], 6) : null;
        $station = isset($cells[1]) ? trim(preg_replace('/\s+/', ' ', (string)$cells[1])) : '';

        $row = array_merge(
            [$river, $rowYear, MONTHS[$rowMonth - 1], (int)$a, $station, $lat, $lon,
             $dt ? $dt->format('Y-m-d') : '', parseTime($cells[5] ?? null)],
            $values,
             [implode('; ', $flags)]
        );

        $sig = implode('|', array_map(fn($x) => (string)$x, $row));
        if (isset($seen[$sig])) { $stats['duplicates']++; continue; }
        $seen[$sig] = true;
        $clean[] = $row;
    }
    return [$clean, $stats];
}

/* "JANUARY" -> 1 ... "DECEMBER" -> 12 */
function monthNum(string $name): int {
    return (int)array_search($name, MONTHS, true) + 1;
}
 
/* Median of a list of numbers (average of the two middle values if the count is even). */
function medianOf(array $v): ?float {
    if (!$v) return null;
    sort($v);
    $n = count($v);
    $m = intdiv($n, 2);
    return ($n % 2) ? (float)$v[$m] : ($v[$m - 1] + $v[$m]) / 2;
}
 
/* Mode = most frequent value. Ties are broken by taking the smallest value so results are repeatable. */
function modeOf(array $v) {
    if (!$v) return null;
    $counts = array_count_values($v);
    $top = array_keys($counts, max($counts));
    sort($top);
    return $top[0];
}
 
/* Fills ONE missing cell.
   $series    = this station's real readings for this parameter, oldest first: [[monthIndex, monthNo, value], ...]
   $t         = month index of the missing cell (year*12 + month)
   $mo        = calendar month (1-12) of the missing cell
   $riverVals = every real reading of this parameter in this river
   $allVals   = every real reading of this parameter in all rivers
   Returns [value, 'MEAN'|'MEDIAN'] or [null, null] if nothing could be found. */
function imputeValue(array $series, int $t, int $mo, array $riverVals, array $allVals): array {
    /* nearest real reading BEFORE ($prev) and AFTER ($next) this month at this station */
    $prev = $next = null;
    foreach ($series as $o) {
        if ($o[0] < $t) $prev = $o;                       // keeps the latest one before $t
        elseif ($o[0] > $t) { $next = $o; break; }        // first one after $t
    }
 
    /* RULE 1: preceding + succeeding values exist and the gap is short -> MEAN */
    if ($prev && $next && ($next[0] - $prev[0] - 1) <= MAX_GAP_FOR_MEAN) {
        return [($prev[2] + $next[2]) / 2, 'MEAN'];
    }
 
    /* RULE 2: large gap -> MEDIAN of the same calendar month in other years (same station) */
    $sameMonth = [];
    foreach ($series as $o) if ($o[1] === $mo) $sameMonth[] = $o[2];
    if (count($sameMonth) >= MIN_SAME_MONTH) return [medianOf($sameMonth), 'MEDIAN'];
 
    /* RULE 3: MEDIAN of everything this station ever recorded for this parameter */
    if ($series) return [medianOf(array_column($series, 2)), 'MEDIAN'];
 
    /* RULE 4: MEDIAN of the whole river;  RULE 5: MEDIAN of all rivers */
    if ($riverVals) return [medianOf($riverVals), 'MEDIAN'];
    if ($allVals)   return [medianOf($allVals),   'MEDIAN'];
 
    return [null, null];
}
 
/* Imputes every blank parameter, Date and Time in every river.
   Input : [river => rows]            (rows still contain null / '' blanks)
   Output: [[river => completed rows], [river => ['imputed'=>, 'dates'=>, 'times'=>]]] */
function imputeRivers(array $byRiver): array {
    global $headers;
    $names = array_slice($headers, COL_P0, N_PARAMS);   // parameter names, used in Imputed_Flags
 
    /* ---- STEP 1: collect the REAL (observed) values, before anything is filled ---- */
    $series = [];     // [river][station][param] = list of [monthIndex, monthNo, value]
    $riverVals = [];  // [river][param] = all real values
    $allVals = [];    // [param] = all real values (all rivers)
    $dayObs = [];     // [river]['ym'|'mo'|'all'] = real days-of-month seen
    $timeObs = [];    // [river]['stmo'|'st'|'all'] = real times seen
 
    foreach ($byRiver as $river => $rows) {
        foreach ($rows as $r) {
            $st = $r[COL_STATION];
            $mo = monthNum($r[COL_MONTH]);
            $ym = $r[COL_YEAR] * 12 + $mo;               // month index: puts all months on one timeline
 
            for ($p = 0; $p < N_PARAMS; $p++) {
                $v = $r[COL_P0 + $p];
                if ($v === null) continue;                // blank -> not a donor
                $series[$river][$st][$p][] = [$ym, $mo, $v];
                $riverVals[$river][$p][] = $v;
                $allVals[$p][] = $v;
            }
            if ($r[COL_DATE] !== '') {                    // day-of-month donors for the Date mode
                $day = (int)substr($r[COL_DATE], 8, 2);
                $dayObs[$river]['ym'][$ym][] = $day;      // same river, same year+month (other stations)
                $dayObs[$river]['mo'][$mo][] = $day;      // same river, same calendar month, other years
                $dayObs[$river]['all'][]     = $day;      // whole river
            }
            if ($r[COL_TIME] !== '') {                    // time donors for the Time mode
                $timeObs[$river]['stmo'][$st][$mo][] = $r[COL_TIME];   // same station + calendar month
                $timeObs[$river]['st'][$st][]        = $r[COL_TIME];   // same station
                $timeObs[$river]['all'][]            = $r[COL_TIME];   // whole river
            }
        }
    }
    /* make sure every station/parameter series runs oldest -> newest (needed for prev/next lookup) */
    foreach ($series as $rv => $byStation) {
        foreach ($byStation as $st => $byParam) {
            foreach ($byParam as $p => $list) {
                usort($list, fn($a, $b) => $a[0] <=> $b[0]);
                $series[$rv][$st][$p] = $list;
            }
        }
    }
 
    /* ---- STEP 2: fill every blank ---- */
    $out = [];
    $impStats = [];
    foreach ($byRiver as $river => $rows) {
        $s = ['imputed' => 0, 'dates' => 0, 'times' => 0];
 
        foreach ($rows as $k => $r) {
            $st = $r[COL_STATION];
            $mo = monthNum($r[COL_MONTH]);
            $ym = $r[COL_YEAR] * 12 + $mo;
            $flags = [];
 
            /* was this a "No Sampling Conducted" row? (all 12 parameters blank) */
            $noData = true;
            for ($p = 0; $p < N_PARAMS; $p++) if ($r[COL_P0 + $p] !== null) { $noData = false; break; }
 
            /* 2a. parameters -> MEAN / MEDIAN */
            for ($p = 0; $p < N_PARAMS; $p++) {
                if ($r[COL_P0 + $p] !== null) continue;
                [$val, $how] = imputeValue(
                    $series[$river][$st][$p] ?? [], $ym, $mo,
                    $riverVals[$river][$p] ?? [], $allVals[$p] ?? []
                );
                if ($val === null) continue;
                $r[COL_P0 + $p] = in_array($p, COLIFORM_PARAMS, true) ? round($val) : round($val, 2);
                $s['imputed']++;
            }
 
            /* 2b. Date -> MODE of the day-of-month, then rebuilt as YYYY-MM-DD */
            if ($r[COL_DATE] === '') {
                $day = modeOf($dayObs[$river]['ym'][$ym] ?? [])      // 1) other stations, same month (same sampling trip)
                    ?? modeOf($dayObs[$river]['mo'][$mo] ?? [])      // 2) same calendar month in other years (2012 case)
                    ?? modeOf($dayObs[$river]['all'] ?? []);         // 3) any date in this river
                if ($day !== null) {
                    $daysInMonth = (int)(new DateTime(sprintf('%04d-%02d-01', $r[COL_YEAR], $mo)))->format('t');
                    $r[COL_DATE] = sprintf('%04d-%02d-%02d', $r[COL_YEAR], $mo, min((int)$day, $daysInMonth)); // e.g. day 31 in Feb -> 28
                    $s['dates']++;
                }
            }
 
            /* 2c. Time -> MODE of the HH:MM values */
            if ($r[COL_TIME] === '') {
                $tm = modeOf($timeObs[$river]['stmo'][$st][$mo] ?? [])  // 1) same station, same calendar month
                   ?? modeOf($timeObs[$river]['st'][$st] ?? [])         // 2) same station, any month
                   ?? modeOf($timeObs[$river]['all'] ?? []);            // 3) any time in this river
                if ($tm !== null) {
                    $r[COL_TIME] = $tm;
                    $s['times']++;
                }
            }
            $rows[$k] = $r;
        }
        $out[$river] = $rows;
        $impStats[$river] = $s;
    }
    return [$out, $impStats];
}

function writeCsv(string $path, array $headers, array $rows): void {
    $h = fopen($path, 'w');
    fputcsv($h, $headers, ',', '"', '\\');
    foreach ($rows as $r) fputcsv($h, $r, ',', '"', '\\');
    fclose($h);
}

function xlsxCol(int $i): string {
    $s = '';
    for ($i++; $i > 0; $i = intdiv($i - 1, 26)) $s = chr(65 + ($i - 1) % 26) . $s;
    return $s;
}

function writeXlsx(string $path, array $headers, array $rows, int $dateCol = 7): void {
    $esc = fn($v) => htmlspecialchars((string)$v, ENT_XML1 | ENT_QUOTES, 'UTF-8');
    $X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

    $cols = '';
    foreach ($headers as $i => $h) {
        $w = $h === 'Station_Name' ? 48 : (in_array($h, ['Date', 'Qualifier_Flags', 'Fecal_Coliform_MPN', 'Total_Coliform_MPN']) ? 20 : 14);
        $cols .= '<col min="' . ($i + 1) . '" max="' . ($i + 1) . '" width="' . $w . '" customWidth="1"/>';
    }

    $sheet = $X . '<worksheet xmlns="' . $ns . '"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>' . $cols . '</cols><sheetData>';
    $sheet .= '<row r="1">';
    foreach ($headers as $i => $h) $sheet .= '<c r="' . xlsxCol($i) . '1" t="inlineStr" s="2"><is><t>' . $esc($h) . '</t></is></c>';
    $sheet .= '</row>';

    foreach ($rows as $r => $row) {
        $n = $r + 2;
        $sheet .= '<row r="' . $n . '">';
        foreach ($row as $i => $v) {
            if ($v === null || $v === '') continue;
            $ref = xlsxCol($i) . $n;
            if ($i === $dateCol && preg_match('/^\d{4}-\d{2}-\d{2}$/', (string)$v)) {
                $serial = (int)(strtotime($v . ' UTC') / 86400) + 25569;
                $sheet .= '<c r="' . $ref . '" s="1"><v>' . $serial . '</v></c>';
            } elseif (is_int($v) || is_float($v)) {
                $sheet .= '<c r="' . $ref . '"><v>' . $v . '</v></c>';
            } else {
                $sheet .= '<c r="' . $ref . '" t="inlineStr"><is><t>' . $esc($v) . '</t></is></c>';
            }
        }
        $sheet .= '</row>';
    }
    $sheet .= '</sheetData></worksheet>';

    $styles = $X . '<styleSheet xmlns="' . $ns . '"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\-mm\-dd"/></numFmts>'
        . '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        . '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        . '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        . '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        . '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        . '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>';

    $zip = new ZipArchive();
    $zip->open($path, ZipArchive::CREATE | ZipArchive::OVERWRITE);
    $zip->addFromString('[Content_Types].xml', $X . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
    $zip->addFromString('_rels/.rels', $X . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    $zip->addFromString('xl/workbook.xml', $X . '<workbook xmlns="' . $ns . '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Cleaned" sheetId="1" r:id="rId1"/></sheets></workbook>');
    $zip->addFromString('xl/_rels/workbook.xml.rels', $X . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    $zip->addFromString('xl/styles.xml', $styles);
    $zip->addFromString('xl/worksheets/sheet1.xml', $sheet);
    $zip->close();
}

/* --------------------------------run pipeline-------------------------------- */

$error = null;
$datasets = [];     // label => ['rows'=>, 'stats'=>, 'file'=>]
$cleanDir = __DIR__ . '/cleaned';

if (!extension_loaded('zip') || !function_exists('simplexml_load_string')) {
    $error = 'This script needs the PHP "zip" and "simplexml" extensions (enable extension=zip in php.ini).';
} elseif (!file_exists($workbookFile)) {
    $error = 'Workbook not found. Place MMORS_water_quality_results_2012-2018_orig.xlsx next to preprocess.php.';
} else {
    try {
        if (!is_dir($cleanDir)) mkdir($cleanDir, 0775, true);
        $sheets = readWorkbook($workbookFile);
        $byRiver = [];     // river => rows
        $rawStats = [];    // river => stats from cleaning
        foreach ($rivers as $river) {
            if (!isset($sheets[$river])) continue;
            [$byRiver[$river], $rawStats[$river]] = processSheet($sheets[$river], $river);
        }
 
        /* 2) IMPUTE all blanks (needs all rivers loaded at once because the last-resort fallback is cross-river) */
        [$byRiver, $impStats] = imputeRivers($byRiver);
 
        /* 3) WRITE the files and build the dataset list */
        $allRows = [];
        $allStats = ['initial' => 0, 'missing' => 0, 'duplicates' => 0, 'corrected' => 0, 'flagged' => 0,
                     'imputed' => 0, 'dates' => 0, 'times' => 0];     // CHANGED: 3 new counters
 
        foreach ($byRiver as $river => $rows) {
            $stats = array_merge($rawStats[$river], $impStats[$river]);
            $file = "$cleanDir/MMORS_{$river}_cleaned.csv";
            writeCsv($file, $headers, $rows);
            writeXlsx("$cleanDir/MMORS_{$river}_cleaned.xlsx", $headers, $rows);
            $datasets[$river] = ['rows' => $rows, 'stats' => $stats, 'file' => $file];
            $allRows = array_merge($allRows, $rows);
            foreach ($allStats as $k => $_) $allStats[$k] += $stats[$k];
        }
        if ($datasets) {
            $file = "$cleanDir/MMORS_All_Rivers_cleaned.csv";
            writeCsv($file, $headers, $allRows);
            writeXlsx("$cleanDir/MMORS_All_Rivers_cleaned.xlsx", $headers, $allRows);
            $datasets[$combinedLabel] = ['rows' => $allRows, 'stats' => $allStats, 'file' => $file];
        }
    } catch (Throwable $e) {
        $error = 'Could not process the workbook: ' . $e->getMessage();
    }
}

$selected = $_GET['dataset'] ?? array_key_first($datasets) ?? $rivers[0];
if (!isset($datasets[$selected]) && $datasets) $selected = array_key_first($datasets);

/* Download (whitelisted by dataset label, served only from ./cleaned) */
if (isset($_GET['download']) && isset($datasets[$_GET['download']])) {
    $f = $datasets[$_GET['download']]['file'];
    $isXlsx = ($_GET['format'] ?? 'xlsx') === 'xlsx';
    if ($isXlsx) $f = preg_replace('/\.csv$/', '.xlsx', $f);
    if (file_exists($f)) {
        header('Content-Type: ' . ($isXlsx
            ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            : 'text/csv'));
        header('Content-Disposition: attachment; filename="' . basename($f) . '"');
        header('Content-Length: ' . filesize($f));
        readfile($f);
        exit;
    }
    http_response_code(404);
    exit('File not found: ' . basename($f) . ' (check that writeXlsx() ran and ./cleaned is writable)');
}

$cur = $datasets[$selected] ?? null;
$totalRows = $cur ? count($cur['rows']) : 0;
$totalPages = max(1, (int)ceil($totalRows / $perPage));
$page = min(max(1, (int)($_GET['page'] ?? 1)), $totalPages);
$previewRows = $cur ? array_slice($cur['rows'], ($page - 1) * $perPage, $perPage) : [];
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMORS Water Quality Analytics</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <style>
        /* Pastel Shades of Blue, Gray, White, and Black */
        body { 
            background-color: #f1f5f9; 
            color: #1e293b; 
            font-family: 'Inter', system-ui, -apple-system, sans-serif; 
            min-height: 100vh;
        }
        
        .report-header {
            background-color: #dbeafe;
            color: #1e3a8a;
            border-radius: 16px;
            padding: 30px 35px;
            margin-bottom: 25px;
            border: 1px solid #93c5fd;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.04);
        }

        .card { 
            background-color: #ffffff; 
            border: 1px solid #cbd5e1; 
            border-radius: 16px !important; 
            overflow: hidden; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.04); 
            color: #1e293b; 
            margin-bottom: 25px;
        }

        .table-container { 
            max-height: 550px; 
            overflow: auto; 
            border-radius: 0px; 
            border: 1px solid #cbd5e1; 
            background-color: #ffffff; 
        }
        .table-container td { 
            white-space: nowrap; 
            color: #1e293b !important; 
            font-weight: 500;
            font-size: 1rem; 
        }
        .table-hover tbody tr:hover { 
            background-color: #f8fafc !important; 
        }
        thead.sticky-top th { 
            background-color: #1e293b !important; 
            color: #ffffff !important; 
            position: sticky; 
            top: 0; 
            z-index: 10; 
            font-weight: 700; 
            font-size: 1rem;
            border-bottom: 2px solid #3b82f6; 
        }
        
        .river-select-box {
            background-color: #ffffff;
            border: 1px solid #94a3b8;
            color: #1e293b;
            font-weight: 600;
            font-size: 1.15rem;
            padding: 12px 20px;
            border-radius: 12px; 
            width: 100%;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.03);
            transition: all 0.2s ease;
        }
        .river-select-box:focus {
            background-color: #ffffff;
            color: #1e293b;
            border-color: #3b82f6;
            box-shadow: 0 0 0 0.2rem rgba(59, 130, 246, 0.25);
        }

        .technical-summary-panel {
            background-color: #eff6ff; 
            border-radius: 16px; 
            border: 1px solid #bfdbfe;
            box-shadow: 0 4px 15px rgba(0,0,0,0.04);
            color: #1e293b;
            padding: 25px 30px;
            margin-bottom: 25px;
        }

        .fab-download {
            position: fixed;
            bottom: 25px;
            right: 25px;
            z-index: 1050;
            border-radius: 50rem;
            padding: 14px 28px;
            font-size: 1.1rem;
            font-weight: 700;
            background-color: #1e293b; 
            color: #ffffff;
            border: 1px solid #475569;
            box-shadow: 0 8px 25px rgba(30, 41, 59, 0.25);
            transition: all 0.2s ease;
        }
        .fab-download:hover {
            transform: translateY(-3px);
            background-color: #0f172a;
            box-shadow: 0 12px 30px rgba(30, 41, 59, 0.35);
            color: #ffffff;
        }
    </style>
</head>
<body>
<div class="container-fluid px-4 py-4" style="max-width: 1500px;">
    <!-- Header with Centered Title "DATA PREPROCESSING" and Right-Aligned Info with Vertical Divider -->
    <div class="report-header">
        <div class="row align-items-center g-3">
            <div class="col-lg-7 text-center text-lg-start">
                <h1 class="fw-bold display-6 mb-0" style="color: #1e3a8a; letter-spacing: 0.5px;">DATA PREPROCESSING</h1>
            </div>
            <div class="col-lg-5 text-center text-lg-end border-start-lg border-blue-300 ps-lg-4">
                <div class="mb-1">
                    <span class="badge px-3 py-2 rounded-pill fw-bold fs-6" style="background-color: #bfdbfe; color: #1e3a8a; border: 1px solid #60a5fa;">
                        Programming Language Used: PHP
                    </span>
                </div>
                <div>
                    <span class="fs-6 fw-semibold" style="color: #334155;">Dataset Range: 2012–2018</span>
                </div>
            </div>
        </div>
    </div>

    <?php if ($error): ?>
        <div class="alert alert-danger bg-danger text-white fs-5 rounded-4"><i class="fa-solid fa-triangle-exclamation me-2"></i><?= htmlspecialchars($error) ?></div>
    <?php endif; ?>

    <!-- Interactive River Dropdown Selector (Stacked & Cleaned Spacing) -->
    <div class="card p-4 shadow-sm">
        <div class="row align-items-center g-3">
            <div class="col-12">
                <label for="riverSelect" class="form-label fw-bold mb-2 fs-5" style="color: #1e293b;">
                    <i class="fa-solid fa-filter me-2 text-primary"></i>River / Dataset Selection:
                </label>
                <select id="riverSelect" class="form-select river-select-box" onchange="location = this.value;">
                    <?php foreach ($datasets as $label => $d): ?>
                        <option value="?dataset=<?= urlencode($label) ?>" <?= ($selected === $label) ? 'selected' : '' ?>>
                            <?= htmlspecialchars($label) ?>
                        </option>
                    <?php endforeach; ?>
                </select>
            </div>
        </div>
    </div>

    <?php if ($cur): $s = $cur['stats']; ?>
        <!-- Technical Summary Panel -->
        <div class="technical-summary-panel shadow-sm">
            <div class="text-uppercase fw-bold border-bottom border-blue-200 pb-2 mb-3 fs-5" style="letter-spacing: 0.5px; color: #1e3a8a;">
                <i class="fa-solid fa-terminal me-2"></i>Data Cleansing Summary &mdash; [<?= htmlspecialchars($selected) ?>]
            </div>
            <div class="row g-3" style="font-family: monospace; font-size: 1rem; color: #334155;">
                <div class="col-md-4">
                    <div class="mb-2">&bull; Initial Records Read:<br><strong class="text-dark fs-4"><?= number_format($s['initial']) ?></strong></div>
                </div>
                <div class="col-md-4">
                    <div class="mb-2">&bull; Final Cleaned Rows:<br><strong class="fs-4" style="color: #2563eb;"><?= number_format(count($cur['rows'])) ?></strong></div>
                </div>
                <div class="col-md-4">
                    <div class="mb-2">&bull; No Samping Rows Restored:<br><strong class="fs-4" style="color: #d97706;"><?= number_format($s['missing']) ?></strong></div>
                </div>
                <div class="col-md-4">
                    <div class="mb-0">&bull; Duplicates Dropped:<br><strong class="fs-4" style="color: #dc2626;"><?= number_format($s['duplicates']) ?></strong></div>
                </div>
                <div class="col-md-4">
                    <div class="mb-0">&bull; Dates Corrected:<br><strong class="fs-4" style="color: #0284c7;"><?= number_format($s['corrected']) ?></strong></div>
                </div>
                <div class="col-md-4">
                    <div class="mb-0">&bull; Flags Generated:<br><strong class="text-dark fs-4"><?= number_format($s['flagged']) ?></strong></div>
                </div>
                  <div class="col-md-4">
                    <div class="mb-0">&bull; Values Imputed (Mean/Median):<br><strong class="fs-4" style="color: #7c3aed;"><?= number_format($s['imputed']) ?></strong></div>
                </div>
                <div class="col-md-4">
                    <div class="mb-0">&bull; Dates Imputed (Mode):<br><strong class="fs-4" style="color: #7c3aed;"><?= number_format($s['dates']) ?></strong></div>
                </div>
                <div class="col-md-4">
                    <div class="mb-0">&bull; Times Imputed (Mode):<br><strong class="fs-4" style="color: #7c3aed;"><?= number_format($s['times']) ?></strong></div>
                </div>
            </div>
        </div>
        
        <div id="dataset-view">
            <div class="card shadow-sm">
                <div class="card-header py-3 px-4 d-flex justify-content-between align-items-center border-bottom border-slate-300 rounded-0" style="background-color: #e2e8f0;">
                    <h5 class="mb-0 fw-bold fs-4" style="color: #1e293b;">
                        <i class="fa-solid fa-table-cells me-2 text-primary"></i>Processed Dataset View:
                        <span style="color: #2563eb;"><?= htmlspecialchars($selected) ?></span>
                    </h5>
                    <span class="badge bg-secondary px-3 py-2 fs-6 rounded-0">
                        Rows <?= $totalRows ? (($page - 1) * $perPage + 1) : 0 ?>–<?= min($page * $perPage, $totalRows) ?> of <?= number_format($totalRows) ?>
                    </span>
                </div>
                <div class="card-body p-0">
                    <?php if ($previewRows): ?>
                        <div class="table-container">
                            <table class="table table-hover align-middle mb-0">
                                <thead class="sticky-top">
                                    <tr><?php foreach ($headers as $head): ?><th class="py-3 px-3"><?= htmlspecialchars($head) ?></th><?php endforeach; ?></tr>
                                </thead>
                                <tbody>
                                    <?php foreach ($previewRows as $row): ?>
                                        <tr><?php foreach ($row as $cell): ?><td class="px-3"><?= htmlspecialchars((string)$cell) ?></td><?php endforeach; ?></tr>
                                    <?php endforeach; ?>
                                </tbody>
                            </table>
                        </div>
                        <div class="d-flex justify-content-between align-items-center p-3 border-top border-slate-300 rounded-0" style="background-color: #e2e8f0;">
                            <a class="btn btn-outline-dark btn-lg js-page <?= $page <= 1 ? 'disabled' : '' ?>"
                            href="?dataset=<?= urlencode($selected) ?>&page=<?= $page - 1 ?>">&laquo; Previous</a>

                            <form id="pageForm" method="get" class="d-flex align-items-center gap-2 fs-5 mb-0" style="color: #1e293b;">
                                <input type="hidden" name="dataset" value="<?= htmlspecialchars($selected) ?>">
                                <span>Page</span>
                                <input type="number" id="pageInput" name="page" min="1" max="<?= $totalPages ?>"
                                    value="<?= $page ?>" class="form-control form-control-lg bg-white text-dark border-slate-400 text-center fw-bold rounded-0" style="width:80px">
                                <span>of <?= $totalPages ?></span>
                            </form>

                            <a class="btn btn-outline-dark btn-lg js-page <?= $page >= $totalPages ? 'disabled' : '' ?>"
                            href="?dataset=<?= urlencode($selected) ?>&page=<?= $page + 1 ?>">Next &raquo;</a>
                        </div>
                    <?php else: ?>
                        <div class="text-center py-5">
                            <i class="fa-solid fa-folder-open text-muted fa-4x mb-3"></i>
                            <p class="text-muted fs-5 mb-0">No records found for this dataset selection.</p>
                        </div>
                    <?php endif; ?>
                </div>
            </div>
        </div>

        <!-- Floating Download Button -->
        <button type="button" class="btn btn-lg fab-download" data-bs-toggle="modal" data-bs-target="#downloadConfirmModal">
            <i class="fa-solid fa-file-excel me-2"></i>Download
        </button>

        <!-- Download Confirmation Modal -->
        <div class="modal fade" id="downloadConfirmModal" tabindex="-1" aria-labelledby="downloadConfirmModalLabel" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered">
                <div class="modal-content bg-white text-dark border border-slate-400 shadow-lg rounded-4">
                    <div class="modal-header border-0 pb-0">
                        <h5 class="modal-title fw-bold fs-4 text-dark" id="downloadConfirmModalLabel">
                            <i class="fa-solid fa-file-excel text-primary me-2"></i>Confirm Download
                        </h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body py-4">
                        <p class="text-secondary mb-0 fs-5">
                            Are you sure you want to download <strong class="text-dark">"<?= htmlspecialchars($selected) ?>"</strong> file?
                        </p>
                    </div>
                    <div class="modal-footer border-0 pt-0">
                        <button type="button" class="btn btn-outline-secondary px-4 py-2 fs-5 rounded-pill" data-bs-dismiss="modal">Cancel</button>
                        <a href="?dataset=<?= urlencode($selected) ?>&download=<?= urlencode($selected) ?>&format=xlsx" class="btn btn-dark px-4 py-2 fs-5 rounded-pill shadow-sm">
                            Yes, Download
                        </a>
                    </div>
                </div>
            </div>
        </div>
    <?php endif; ?>
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
    async function loadView(url, push = true) {
    const y = window.scrollY;
    try {
        const res = await fetch(url, { headers: { 'X-Requested-With': 'fetch' } });
        const doc = new DOMParser().parseFromString(await res.text(), 'text/html');
        const fresh = doc.getElementById('dataset-view');
        if (!fresh) { window.location.href = url; return; }
        document.getElementById('dataset-view').replaceWith(fresh);
        if (push) history.pushState(null, '', url);
        window.scrollTo(0, y);              // stay exactly where you were
    } catch (e) { window.location.href = url; }
    }

    // Previous / Next
    document.addEventListener('click', e => {
        const a = e.target.closest('a.js-page');
        if (!a) return;
        e.preventDefault();
        if (!a.classList.contains('disabled')) loadView(a.href);
    });

    // Typed page number (Enter or leaving the field)
    function goToTypedPage(input) {
        const max = parseInt(input.max) || 1;
        const v = Math.min(Math.max(1, parseInt(input.value) || 1), max);
        input.value = v;
        const p = new URLSearchParams(new FormData(input.form));
        p.set('page', v);
        loadView('?' + p.toString());
    }
    document.addEventListener('change', e => {
        if (e.target.id === 'pageInput') goToTypedPage(e.target);
    });
    document.addEventListener('submit', e => {
        if (e.target.id === 'pageForm') {
            e.preventDefault();
            goToTypedPage(e.target.querySelector('#pageInput'));
        }
    });

    // Browser back/forward
    window.addEventListener('popstate', () => loadView(location.href, false));
    </script>
</body>
</html>