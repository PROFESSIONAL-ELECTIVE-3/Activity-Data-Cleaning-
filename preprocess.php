<?php
/**
 * MMORS Water Quality Data Preprocessing - Web UI Viewer
 * Group Activity: Data Preprocessing Using PHP
 */

$csvFiles = [
    'MMORS_water_quality_results_2012-2018_orig-Marilao.csv',
    'MMORS_water_quality_results_2012-2018_orig-Meycauayan.csv',
    'MMORS_water_quality_results_2012-2018_orig-Obando.csv',
    'MMORS_water_quality_results_2012-2018_orig-Table1 (3).csv'
];

$processingResults = [];

foreach ($csvFiles as $inputFile) {
    if (!file_exists($inputFile)) continue;

    $outputFile = str_replace('.csv', '_cleaned.csv', $inputFile);
    $rawRows = [];
    $headers = [];
    
    // Load and clean
    if (($handle = fopen($inputFile, 'r')) !== FALSE) {
        $rowIndex = 0;
        while (($row = fgetcsv($handle, 10000, ',')) !== FALSE) {
            if ($rowIndex == 7 || $rowIndex == 11) {
                $headers = $row;
            } elseif ($rowIndex > 11 && !empty($headers)) {
                if (count($row) === count($headers)) {
                    $rawRows[] = array_combine($headers, $row);
                }
            }
            $rowIndex++;
        }
        fclose($handle);
    }

    if (empty($headers)) {
        $handle = fopen($inputFile, 'r');
        $headers = fgetcsv($handle, 10000, ',');
        while (($row = fgetcsv($handle, 10000, ',')) !== FALSE) {
            if (count($row) === count($headers)) {
                $rawRows[] = array_combine($headers, $row);
            }
        }
        fclose($handle);
    }

    $initialCount = count($rawRows);
    $cleanedData = [];
    $duplicateTracker = [];
    $missingCount = 0;
    $duplicatesRemoved = 0;

    foreach ($rawRows as $row) {
        $keys = array_keys($row);
        $firstCol = $keys[0] ?? null;
        if ($firstCol === null || empty(trim($row[$firstCol]))) {
            $missingCount++;
            continue;
        }

        foreach ($row as $key => $value) {
            $row[$key] = trim($value);
        }

        $rowSignature = implode('|', $row);
        if (isset($duplicateTracker[$rowSignature])) {
            $duplicatesRemoved++;
            continue;
        }
        $duplicateTracker[$rowSignature] = true;
        $cleanedData[] = $row;
    }

    $finalCount = count($cleanedData);

    // Save cleaned file
    $outputHandle = fopen($outputFile, 'w');
    if (!empty($cleanedData) && !empty($headers)) {
        fputcsv($outputHandle, $headers);
        foreach ($cleanedData as $row) {
            fputcsv($outputHandle, $row);
        }
    }
    fclose($outputHandle);

    // Store metrics for HTML display
    $processingResults[] = [
        'file' => $inputFile,
        'cleaned_file' => $outputFile,
        'initial' => $initialCount,
        'final' => $finalCount,
        'missing' => $missingCount,
        'duplicates' => $duplicatesRemoved
    ];
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMORS Water Quality Preprocessing Results</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
    <div class="container py-5">
        <header class="mb-4 text-center">
            <h1 class="fw-bold text-primary">MMORS Water Quality Results</h1>
            <p class="text-muted">Data Preprocessing Dashboard (PHP & HTML)</p>
        </header>

        <div class="row g-4">
            <?php foreach ($processingResults as $res): ?>
                <div class="col-md-6">
                    <div class="card shadow-sm border-0">
                        <div class="card-header bg-dark text-white">
                            <h5 class="card-title mb-0 fs-6 text-truncate"><?= htmlspecialchars($res['file']) ?></h5>
                        </div>
                        <div class="card-body">
                            <ul class="list-group list-group-flush mb-3">
                                <li class="list-group-item d-flex justify-content-between align-items-center">
                                    Initial Records <span class="badge bg-secondary rounded-pill"><?= $res['initial'] ?></span>
                                </li>
                                <li class="list-group-item d-flex justify-content-between align-items-center">
                                    Final Cleaned Records <span class="badge bg-success rounded-pill"><?= $res['final'] ?></span>
                                </li>
                                <li class="list-group-item d-flex justify-content-between align-items-center">
                                    Missing Core Data Skipped <span class="badge bg-warning text-dark rounded-pill"><?= $res['missing'] ?></span>
                                </li>
                                <li class="list-group-item d-flex justify-content-between align-items-center">
                                    Duplicates Removed <span class="badge bg-danger rounded-pill"><?= $res['duplicates'] ?></span>
                                </li>
                            </ul>
                            <div class="alert alert-info py-2 mb-0 small">
                                <strong>Output Saved:</strong> <code><?= htmlspecialchars($res['cleaned_file']) ?></code>
                            </div>
                        </div>
                    </div>
                </div>
            <?php endforeach; ?>
        </div>
    </div>
</body>
</html>