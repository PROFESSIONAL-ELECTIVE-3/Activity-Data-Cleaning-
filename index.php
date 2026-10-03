<?php
/**
 * MMORS Water Quality Data Preprocessing Dashboard
 * Group Activity: Data Preprocessing Using PHP
 *
 * Reads the ORIGINAL workbook (all 18 columns) directly, so nothing important is lost:
 * station name, coordinates, date, time, and all 12 parameters (incl. BOD, Total Coliform, Ammonia).
 * The workbook is never modified; cleaned CSVs are written to ./cleaned/
 *
 * Requires: PHP ZipArchive + SimpleXML (enabled by default in XAMPP; php-zip / php-xml on Linux).
 * Put MMORS_water_quality_results_2012-2018_orig.xlsx in the same folder as this file.
 */

$workbookFile = __DIR__ . '/MMORS_water_quality_results_2012-2018_orig.xlsx';
$rivers = ['Marilao', 'Meycauayan', 'Obando'];            // sheet names in the workbook
$combinedLabel = 'All Rivers';
$perPage = 25;

// Workbook columns G..R (index 6..17) -> output column names
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

/* ---------------------------------------------------------------- XLSX reader */

function colIndex(string $ref): int
{
    preg_match('/^([A-Z]+)/', $ref, $m);
    $n = 0;
    foreach (str_split($m[1]) as $ch) $n = $n * 26 + (ord($ch) - 64);
    return $n - 1;
}

/** Returns [sheetName => list of rows], each row = [colIndex => string|float|null]. */
function readWorkbook(string $path): array
{
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
                if ($t === 's')              $val = $shared[(int)$c->v] ?? null;
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

/* ---------------------------------------------------------------- cleaning helpers */

/** Returns [float|null, '<'|'>'|null]. '<0.048*' -> half the limit (flagged); '>16000' -> the bound (flagged). */
function cleanMeasure($v): array
{
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

/** Excel serial number or text like "10/17//2016" -> DateTime|null. */
function parseDate($v): ?DateTime
{
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

/** Excel time fraction or text like "12:06PM" -> "HH:MM" (24h), or ''. */
function parseTime($v): string
{
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

function processSheet(array $rows, string $river): array
{
    global $paramCols;
    $stats = ['initial' => 0, 'missing' => 0, 'duplicates' => 0, 'corrected' => 0, 'flagged' => 0];
    $clean = [];
    $seen = [];
    $year = null;
    $month = null;   // 1-12

    foreach ($rows as $cells) {
        $a = $cells[0] ?? null;

        // Month banner, e.g. "CY 2012 JANUARY" (tolerates typo "CY 20178 DECEMBER")
        if (is_string($a) && preg_match('/^\s*CY\s*(\d{4})\d*\s+([A-Za-z]+)/', $a, $m)) {
            $idx = array_search(strtoupper($m[2]), MONTHS, true);
            if ($idx !== false) { $year = (int)$m[1]; $month = $idx + 1; }
            continue;
        }

        // Station rows start with the station number (1..5); everything else is header/guideline/footnote
        if (!is_float($a) || $a < 1 || floor($a) != $a || $year === null) continue;
        $stats['initial']++;

        // Date column is authoritative when present (banners have a few typos, e.g. "CY 2016 AUGUST" with 2017 dates)
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

        if (!$hasData) {                       // station row with no measurements (e.g. "NO SAMPLING CONDUCTED")
            $stats['missing']++;
            continue;
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

function writeCsv(string $path, array $headers, array $rows): void
{
    $h = fopen($path, 'w');
    fputcsv($h, $headers, ',', '"', '\\');
    foreach ($rows as $r) fputcsv($h, $r, ',', '"', '\\');
    fclose($h);
}

/* ---------------------------------------------------------------- run pipeline */

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
        $allRows = [];
        $allStats = ['initial' => 0, 'missing' => 0, 'duplicates' => 0, 'corrected' => 0, 'flagged' => 0];

        foreach ($rivers as $river) {
            if (!isset($sheets[$river])) continue;
            [$rows, $stats] = processSheet($sheets[$river], $river);
            $file = "$cleanDir/MMORS_{$river}_cleaned.csv";
            writeCsv($file, $headers, $rows);
            $datasets[$river] = ['rows' => $rows, 'stats' => $stats, 'file' => $file];
            $allRows = array_merge($allRows, $rows);
            foreach ($allStats as $k => $_) $allStats[$k] += $stats[$k];
        }
        if ($datasets) {
            $file = "$cleanDir/MMORS_All_Rivers_cleaned.csv";
            writeCsv($file, $headers, $allRows);
            $datasets[$combinedLabel] = ['rows' => $allRows, 'stats' => $allStats, 'file' => $file];
        }
    } catch (Throwable $e) {
        $error = 'Could not process the workbook: ' . $e->getMessage();
    }
}

$selected = $_GET['dataset'] ?? array_key_first($datasets) ?? $rivers[0];
if (!isset($datasets[$selected]) && $datasets) $selected = array_key_first($datasets);

// Download (whitelisted by dataset label, served only from ./cleaned)
if (isset($_GET['download']) && isset($datasets[$_GET['download']])) {
    $f = $datasets[$_GET['download']]['file'];
    if (file_exists($f)) {
        header('Content-Type: text/csv');
        header('Content-Disposition: attachment; filename="' . basename($f) . '"');
        readfile($f);
        exit;
    }
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
        body { background-color: #f8f9fa; font-family: 'Inter', system-ui, -apple-system, sans-serif; }
        .hero-banner { background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: white; border-radius: 16px; padding: 30px; margin-bottom: 30px; }
        .nav-tabs .nav-link { color: #64748b; border: none; font-weight: 600; padding: 12px 24px; border-radius: 8px; transition: all 0.2s ease; }
        .nav-tabs .nav-link:hover { color: #0f172a; background-color: #e2e8f0; }
        .nav-tabs .nav-link.active { color: #fff; background-color: #0d6efd; box-shadow: 0 4px 12px rgba(13, 110, 253, 0.25); }
        .card { border: none; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.03); }
        .metric-card { border-top: 4px solid #0d6efd; }
        .table-container { max-height: 520px; overflow: auto; border-radius: 10px; border: 1px solid #e2e8f0; }
        .table-container td, .table-container th { white-space: nowrap; }
        thead.sticky-top th { background-color: #0f172a !important; color: white; position: sticky; top: 0; z-index: 10; font-weight: 500; letter-spacing: 0.5px; }
    </style>
</head>
<body>
<div class="container-fluid px-4 py-5">
    <div class="hero-banner shadow-sm d-flex flex-wrap justify-content-between align-items-center">
        <div>
            <span class="badge bg-primary bg-opacity-25 text-info px-3 py-2 rounded-pill mb-2 fw-semibold">
                <i class="fa-solid fa-water me-2"></i>PHP Data Preprocessing Pipeline
            </span>
            <h1 class="fw-bold mb-1 display-6">MMORS Water Quality Analytics</h1>
            <p class="text-muted mb-0">Interactive cleaning, verification, and transformation dashboard (2012–2018)</p>
        </div>
        <div class="mt-3 mt-md-0">
            <?php if ($cur): ?>
                <a href="?dataset=<?= urlencode($selected) ?>&download=<?= urlencode($selected) ?>" class="btn btn-success btn-lg px-4 shadow">
                    <i class="fa-solid fa-cloud-arrow-down me-2"></i>Download Cleaned CSV
                </a>
            <?php endif; ?>
        </div>
    </div>

    <?php if ($error): ?>
        <div class="alert alert-danger"><i class="fa-solid fa-triangle-exclamation me-2"></i><?= htmlspecialchars($error) ?></div>
    <?php endif; ?>

    <ul class="nav nav-tabs border-0 gap-2 mb-4">
        <?php foreach ($datasets as $label => $d): ?>
            <li class="nav-item">
                <a class="nav-link <?= ($selected === $label) ? 'active' : '' ?>" href="?dataset=<?= urlencode($label) ?>">
                    <i class="fa-solid fa-chart-line me-2"></i><?= htmlspecialchars($label) ?>
                </a>
            </li>
        <?php endforeach; ?>
    </ul>

    <?php if ($cur): $s = $cur['stats']; ?>
        <div class="row g-4 mb-3">
            <div class="col-md-3"><div class="card metric-card p-4 bg-white">
                <span class="text-muted small text-uppercase fw-bold">Initial Records</span>
                <h2 class="fw-bold text-dark mt-2 mb-0"><?= number_format($s['initial']) ?></h2>
                <span class="badge bg-light text-secondary mt-2 w-50">Station readings</span>
            </div></div>
            <div class="col-md-3"><div class="card metric-card p-4 bg-white" style="border-top-color:#198754;">
                <span class="text-muted small text-uppercase fw-bold">Cleaned Records</span>
                <h2 class="fw-bold text-success mt-2 mb-0"><?= number_format(count($cur['rows'])) ?></h2>
                <span class="badge bg-success bg-opacity-10 text-success mt-2 w-50">Analysis Ready</span>
            </div></div>
            <div class="col-md-3"><div class="card metric-card p-4 bg-white" style="border-top-color:#ffc107;">
                <span class="text-muted small text-uppercase fw-bold">Missing Skipped</span>
                <h2 class="fw-bold text-dark mt-2 mb-0"><?= number_format($s['missing']) ?></h2>
                <span class="badge bg-warning bg-opacity-10 text-dark mt-2 w-50">No measurements</span>
            </div></div>
            <div class="col-md-3"><div class="card metric-card p-4 bg-white" style="border-top-color:#dc3545;">
                <span class="text-muted small text-uppercase fw-bold">Duplicates Dropped</span>
                <h2 class="fw-bold text-danger mt-2 mb-0"><?= number_format($s['duplicates']) ?></h2>
                <span class="badge bg-danger bg-opacity-10 text-danger mt-2 w-50">Redundant rows</span>
            </div></div>
        </div>
        <div class="alert alert-info small">
            <i class="fa-solid fa-circle-info me-2"></i>
            <strong><?= number_format($s['corrected']) ?></strong> row(s) had Year/Month corrected from the Date column (banner typos).
            <strong><?= number_format($s['flagged']) ?></strong> row(s) contain values below the detection limit (stored as half the limit) or above the
            reporting range; see <code>Qualifier_Flags</code>.
        </div>
        <div id="dataset-view">
            <div class="card shadow-sm">
                <div class="card-header bg-white py-3 px-4 d-flex justify-content-between align-items-center border-0">
                    <h5 class="mb-0 fw-bold text-dark">
                        <i class="fa-solid fa-table-cells me-2 text-primary"></i>Processed Dataset View:
                        <span class="text-primary"><?= htmlspecialchars($selected) ?></span>
                    </h5>
                    <span class="badge bg-secondary px-3 py-2">
                        Rows <?= $totalRows ? (($page - 1) * $perPage + 1) : 0 ?>–<?= min($page * $perPage, $totalRows) ?> of <?= number_format($totalRows) ?>
                    </span>
                </div>
                <div class="card-body p-0">
                    <?php if ($previewRows): ?>
                        <div class="table-container">
                            <table class="table table-hover align-middle mb-0 small">
                                <thead class="sticky-top">
                                    <tr><?php foreach ($headers as $head): ?><th class="py-3 px-3"><?= htmlspecialchars($head) ?></th><?php endforeach; ?></tr>
                                </thead>
                                <tbody>
                                    <?php foreach ($previewRows as $row): ?>
                                        <tr><?php foreach ($row as $cell): ?><td class="px-3 text-secondary"><?= htmlspecialchars((string)$cell) ?></td><?php endforeach; ?></tr>
                                    <?php endforeach; ?>
                                </tbody>
                            </table>
                        </div>
                        <div class="d-flex justify-content-between align-items-center p-3">
                            <a class="btn btn-outline-secondary btn-sm js-page <?= $page <= 1 ? 'disabled' : '' ?>"
                            href="?dataset=<?= urlencode($selected) ?>&page=<?= $page - 1 ?>">&laquo; Previous</a>

                            <form id="pageForm" method="get" class="d-flex align-items-center gap-2 small text-muted mb-0">
                                <input type="hidden" name="dataset" value="<?= htmlspecialchars($selected) ?>">
                                <span>Page</span>
                                <input type="number" id="pageInput" name="page" min="1" max="<?= $totalPages ?>"
                                    value="<?= $page ?>" class="form-control form-control-sm text-left" style="width:50px">
                                <span>of <?= $totalPages ?></span>
                            </form>

                            <a class="btn btn-outline-secondary btn-sm js-page <?= $page >= $totalPages ? 'disabled' : '' ?>"
                            href="?dataset=<?= urlencode($selected) ?>&page=<?= $page + 1 ?>">Next &raquo;</a>
                        </div>
                    <?php else: ?>
                        <div class="text-center py-5">
                            <i class="fa-solid fa-folder-open text-muted fa-3x mb-3"></i>
                            <p class="text-muted mb-0">No records found for this dataset selection.</p>
                        </div>
                    <?php endif; ?>
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
        window.scrollTo(0, y);                 // stay exactly where you were
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