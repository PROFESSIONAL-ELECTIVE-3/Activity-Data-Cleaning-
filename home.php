<?php
/**
 * Group Activity: Data Preprocessing Using PHP
 * Dataset: MMORS Water Quality Results (2012-2018)
 * 
 * Instructions for VS Code:
 * 1. Open your project folder in VS Code.
 * 2. Ensure 'mmors_water_quality_raw.csv' is in the same directory.
 * 3. Open the integrated terminal (Ctrl + `) and run: php preprocess.php
 */

$inputFile = 'mmors_water_quality_raw.csv';
$outputFile = 'mmors_water_quality_cleaned.csv';

echo "=== MMORS WATER QUALITY DATA PREPROCESSING ===\n\n";

// Check if input file exists
if (!file_exists($inputFile)) {
    die("❌ Error: The file '{$inputFile}' was not found. Please place it in your workspace directory.\n");
}

echo "📂 Loading dataset from '{$inputFile}'...\n";

// 1. Load and Inspect Dataset
$rawRows = [];
$headers = [];
if (($handle = fopen($inputFile, 'r')) !== FALSE) {
    $headers = fgetcsv($handle, 1000, ',');
    if ($headers === FALSE) {
        die("❌ Error: The CSV file is empty or headers could not be read.\n");
    }
    
    while (($row = fgetcsv($handle, 1000, ',')) !== FALSE) {
        // Ensure row column count matches headers to avoid offset mismatch
        if (count($row) === count($headers)) {
            $rawRows[] = array_combine($headers, $row);
        }
    }
    fclose($handle);
}

$initialCount = count($rawRows);
echo "✔ Successfully loaded {$initialCount} rows.\n\n";
echo "🧹 Executing data cleaning and quality assessment pipeline...\n";

// 2. Data Quality Assessment Tracking
$cleanedData = [];
$duplicateTracker = [];
$metrics = [
    'missing_core_data' => 0,
    'duplicates_removed' => 0,
    'invalid_entries' => 0,
    'outliers_flagged' => 0
];

foreach ($rawRows as $index => $row) {
    // A. Check for missing essential core values (e.g., Station ID and Date)
    if (empty(trim($row['Station_ID'] ?? '')) || empty(trim($row['Date'] ?? ''))) {
        $metrics['missing_core_data']++;
        continue; // Skip records with missing critical identifiers
    }

    // B. Standardize text formatting (Trim whitespace and fix capitalization for names)
    if (isset($row['Station_Name'])) {
        $row['Station_Name'] = trim(ucwords(strtolower($row['Station_Name'])));
    }

    // C. Clean and validate numeric parameters (Example: pH level evaluation)
    if (isset($row['pH'])) {
        $pHVal = trim($row['pH']);
        if (is_numeric($pHVal)) {
            $pHFloat = floatval($pHVal);
            // Environmental validation check: standard pH scale ranges strictly from 0 to 14
            if ($pHFloat < 0 || $pHFloat > 14) {
                $metrics['outliers_flagged']++;
                $row['pH'] = null; // Nullify physical impossibility
            } else {
                $row['pH'] = $pHFloat;
            }
        } else {
            $metrics['invalid_entries']++;
            $row['pH'] = null; // Handle text/invalid characters in numeric columns
        }
    }

    // D. Duplicate Detection using a composite signature (Station ID + Date)
    $rowSignature = trim($row['Station_ID']) . '|' . trim($row['Date']);
    if (isset($duplicateTracker[$rowSignature])) {
        $metrics['duplicates_removed']++;
        continue; // Skip duplicate records
    }
    $duplicateTracker[$rowSignature] = true;

    // Add validated row to the clean collection dataset
    $cleanedData[] = $row;
}

$finalCount = count($cleanedData);

// 3. Export Cleaned Dataset
$outputHandle = fopen($outputFile, 'w');
if (!empty($cleanedData)) {
    fputcsv($outputHandle, $headers); // Write original headers back
    foreach ($cleanedData asphp preprocess.php $row) {
        fputcsv($outputHandle, $row);
    }
}
fclose($outputHandle);

// 4. Print Summary Report (Useful for screenshots, reports, and video presentation)
echo "\n============================================\n";
echo "       PREPROCESSING RESULTS SUMMARY        \n";
echo "============================================\n";
echo "• Initial Total Records   : {$initialCount}\n";
echo "• Final Cleaned Records   : {$finalCount}\n";
echo "• Missing Core Data Skipped: {$metrics['missing_core_data']}\n";
echo "• Duplicate Rows Removed  : {$metrics['duplicates_removed']}\n";
echo "• Invalid Entries Fixed   : {$metrics['invalid_entries']}\n";
echo "• Outliers Flagged/Nullified: {$metrics['outliers_flagged']}\n";
echo "============================================\n";
echo "✔ Cleaned dataset successfully generated and saved to: '{$outputFile}'\n";
?>