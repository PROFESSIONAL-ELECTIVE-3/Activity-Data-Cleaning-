<?php
/**
 * MMORS Water Quality Data Preprocessing Dashboard
 * Group Activity: Data Preprocessing Using PHP
 */

$csvFiles = [
    'Marilao' => 'MMORS_water_quality_results_2012-2018_orig-Marilao.csv',
    'Meycauayan' => 'MMORS_water_quality_results_2012-2018_orig-Meycauayan.csv',
    'Obando' => 'MMORS_water_quality_results_2012-2018_orig-Obando.csv',
    'Table 1' => 'MMORS_water_quality_results_2012-2018_orig-Table1 (3).csv'
];

// Handle CSV Download Request
if (isset($_GET['download'])) {
    $fileToDownload = basename($_GET['download']);
    if (file_exists($fileToDownload)) {
        header('Content-Type: text/csv');
        header('Content-Disposition: attachment; filename="' . $fileToDownload . '"');
        readfile($fileToDownload);
        exit;
    }
}

$dashboardData = [];
$selectedDatasetKey = $_GET['dataset'] ?? array_key_first($csvFiles);
$previewRows = [];
$previewHeaders = [];

foreach ($csvFiles as $label => $inputFile) {
    if (!file_exists($inputFile)) continue;

    $rawRows = [];
    $headers = [];
    
    // Load and clean data directly into memory / overwriting source cleanly
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

    // Overwrite the original file directly with clean data (preventing duplicate _cleaned.csv files)
    $outputHandle = fopen($inputFile, 'w');
    if (!empty($cleanedData) && !empty($headers)) {
        fputcsv($outputHandle, $headers);
        foreach ($cleanedData as $row) {
            fputcsv($outputHandle, $row);
        }
    }
    fclose($outputHandle);

    $dashboardData[$label] = [
        'file' => $inputFile,
        'initial' => $initialCount,
        'final' => $finalCount,
        'missing' => $missingCount,
        'duplicates' => $duplicatesRemoved
    ];

    if ($label === $selectedDatasetKey) {
        $previewHeaders = $headers;
        $previewRows = array_slice($cleanedData, 0, 25); // Preview first 25 rows
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMORS Water Quality Analytics</title>
    <!-- Bootstrap 5 & FontAwesome -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <style>
        body { background-color: #f8f9fa; font-family: 'Inter', system-ui, -apple-system, sans-serif; }
        .hero-banner { background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: white; border-radius: 16px; padding: 30px; margin-bottom: 30px; }
        .nav-tabs .nav-link { color: #64748b; border: none; font-weight: 600; padding: 12px 24px; border-radius: 8px; transition: all 0.2s ease; }
        .nav-tabs .nav-link:hover { color: #0f172a; background-color: #e2e8f0; }
        .nav-tabs .nav-link.active { color: #fff; background-color: #0d6efd; box-shadow: 0 4px 12px rgba(13, 110, 253, 0.25); }
        .card { border: none; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.03); transition: transform 0.2s; }
        .metric-card { border-top: 4px solid #0d6efd; }
        .table-container { max-height: 480px; overflow-y: auto; border-radius: 10px; border: 1px solid #e2e8f0; }
        thead.sticky-top th { background-color: #0f172a !important; color: white; position: sticky; top: 0; z-index: 10; font-weight: 500; letter-spacing: 0.5px; }
    </style>
</head>
<body>

    <div class="container py-5">
        <!-- Hero Header -->
        <div class="hero-banner shadow-sm d-flex flex-wrap justify-content-between align-items-center">
            <div>
                <span class="badge bg-primary bg-opacity-25 text-info px-3 py-2 rounded-pill mb-2 fw-semibold">
                    <i class="fa-solid fa-water me-2"></i>PHP Data Preprocessing Pipeline
                </span>
                <h1 class="fw-bold mb-1 display-6">MMORS Water Quality Analytics</h1>
                <p class="text-muted mb-0">Interactive cleaning, verification, and transformation dashboard (2012–2018)</p>
            </div>
            <div class="mt-3 mt-md-0">
                <?php if (isset($dashboardData[$selectedDatasetKey])): ?>
                    <a href="?dataset=<?= urlencode($selectedDatasetKey) ?>&download=<?= urlencode($dashboardData[$selectedDatasetKey]['file']) ?>" class="btn btn-success btn-lg px-4 shadow">
                        <i class="fa-solid fa-cloud-arrow-down me-2"></i>Download Cleaned CSV
                    </a>
                <?php endif; ?>
            </div>
        </div>

        <!-- Top Tab Navigation -->
        <ul class="nav nav-tabs border-0 gap-2 mb-4">
            <?php foreach ($csvFiles as $label => $file): ?>
                <li class="nav-item">
                    <a class="nav-link <?= ($selectedDatasetKey === $label) ? 'active' : '' ?>" href="?dataset=<?= urlencode($label) ?>">
                        <i class="fa-solid fa-chart-line me-2"></i><?= htmlspecialchars($label) ?>
                    </a>
                </li>
            <?php endforeach; ?>
        </ul>

        <!-- Metric Cards for Selected Dataset -->
        <?php if (isset($dashboardData[$selectedDatasetKey])): 
            $currentRes = $dashboardData[$selectedDatasetKey];
        ?>
            <div class="row g-4 mb-4">
                <div class="col-md-3">
                    <div class="card metric-card p-4 bg-white">
                        <span class="text-muted small text-uppercase fw-bold">Initial Records</span>
                        <h2 class="fw-bold text-dark mt-2 mb-0"><?= number_format($currentRes['initial']) ?></h2>
                        <span class="badge bg-light text-secondary mt-2 w-50">Raw inputs</span>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card metric-card p-4 bg-white" style="border-top-color: #198754;">
                        <span class="text-muted small text-uppercase fw-bold">Cleaned Records</span>
                        <h2 class="fw-bold text-success mt-2 mb-0"><?= number_format($currentRes['final']) ?></h2>
                        <span class="badge bg-success bg-opacity-10 text-success mt-2 w-50">Analysis Ready</span>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card metric-card p-4 bg-white" style="border-top-color: #ffc107;">
                        <span class="text-muted small text-uppercase fw-bold">Missing Skipped</span>
                        <h2 class="fw-bold text-dark mt-2 mb-0"><?= number_format($currentRes['missing']) ?></h2>
                        <span class="badge bg-warning bg-opacity-10 text-dark mt-2 w-50">Null keys dropped</span>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card metric-card p-4 bg-white" style="border-top-color: #dc3545;">
                        <span class="text-muted small text-uppercase fw-bold">Duplicates Dropped</span>
                        <h2 class="fw-bold text-danger mt-2 mb-0"><?= number_format($currentRes['duplicates']) ?></h2>
                        <span class="badge bg-danger bg-opacity-10 text-danger mt-2 w-50">Redundant rows</span>
                    </div>
                </div>
            </div>
        <?php endif; ?>

        <!-- Cleaned Data Table Preview -->
        <div class="card shadow-sm">
            <div class="card-header bg-white py-3 px-4 d-flex justify-content-between align-items-center border-0">
                <div>
                    <h5 class="mb-0 fw-bold text-dark">
                        <i class="fa-solid fa-table-cells me-2 text-primary"></i>Processed Dataset View: <span class="text-primary"><?= htmlspecialchars($selectedDatasetKey) ?></span>
                    </h5>
                </div>
                <span class="badge bg-secondary px-3 py-2">Displaying top <?= count($previewRows) ?> rows</span>
            </div>
            <div class="card-body p-0">
                <?php if (!empty($previewRows)): ?>
                    <div class="table-container">
                        <table class="table table-hover align-middle mb-0 small">
                            <thead class="sticky-top">
                                <tr>
                                    <?php foreach ($previewHeaders as $head): ?>
                                        <th class="py-3 px-3"><?= htmlspecialchars($head) ?></th>
                                    <?php endforeach; ?>
                                </tr>
                            </thead>
                            <tbody>
                                <?php foreach ($previewRows as $row): ?>
                                    <tr>
                                        <?php foreach ($previewHeaders as $head): ?>
                                            <td class="px-3 text-secondary"><?= htmlspecialchars($row[$head] ?? '') ?></td>
                                        <?php endforeach; ?>
                                    </tr>
                                <?php endforeach; ?>
                            </tbody>
                        </table>
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

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>