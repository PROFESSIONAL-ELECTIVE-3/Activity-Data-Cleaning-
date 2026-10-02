// WALA PA
<?php
/**
 * Group Activity: Data Preprocessing Using PHP
 * Dataset: MMORS Water Quality Results (2012-2018)
 * 
 * Instructions:
 * 1. Ensure your 4 CSV files are in the same folder as this script.
 * 2. Run in your VS Code terminal: php preprocess.php
 */

// List of your 4 CSV files
$csvFiles = [
    'MMORS_water_quality_results_2012-2018_orig.csv',
    'MMORS_water_quality_results_2012-2018_orig-Table1 (3).csv',
    'MMORS_water_quality_results_2012-2018_orig-Obando.csv',
    'MMORS_water_quality_results_2012-2018_orig-Meycauayan.csv'
];

echo "====================================================\n";
echo "   MMORS WATER QUALITY DATA PREPROCESSING (PHP)     \n";
echo "====================================================\n\n";

foreach ($csvFiles as $inputFile) {
    // Check if file exists
    if (!file_exists($inputFile)) {
        echo "⚠ Notice: File '{$inputFile}' not found in directory. Skipping...\n\n";
        continue;
    }

    $outputFile = str_replace('.csv', '_cleaned.csv', $inputFile);
    echo "📂 Processing: {$inputFile} ...\n";

    $rawRows = [];
    $headers = [];
    
    // 1. Load and Inspect Dataset
    if (($handle = fopen($inputFile, 'r')) !== FALSE) {
        $rowIndex = 0;
        while (($row = fgetcsv($handle, 1000, ',')) !== FALSE) {
            // Note: The raw MMORS sheets start data headers around row index 11 (after DENR metadata)
            // If your converted CSV has headers on row 0 instead, change 11 to 0.
            if ($rowIndex == 11) {
                $headers = $row;
            } elseif ($rowIndex > 11 && !empty($headers)) {
                if (count($row) === count($headers)) {
                    $rawRows[] = array_combine($headers, $row);
                }
            } elseif ($rowIndex < 11 && count($row) > 2) {
                // Fallback check if headers are at the very top (Row 0)
                // If row 0 contains column names like Station or pH, adjust accordingly.
            }
            $rowIndex++;
        }
        fclose($handle);
    }

    // Fallback if headers weren't captured at row 11 (e.g. if conversion stripped top rows)
    if (empty($headers) && file_exists($inputFile)) {
        $handle = fopen($inputFile, 'r');
        $headers = fgetcsv($handle, 1000, ','); // Grab row 0 as header
        while (($row = fgetcsv($handle, 1000, ',')) !== FALSE) {
            if (count($row) === count($headers)) {
                $rawRows[] = array_combine($headers, $row);
            }
        }
        fclose($handle);
    }

    $initialCount = count($rawRows);
    echo "   ✔ Loaded records: {$initialCount}\n";

    // 2. Data Quality Assessment & Cleaning
    $cleanedData = [];
    $duplicateTracker = [];
    $metrics = [
        'missing_core_data' => 0,
        'duplicates_removed' => 0,
        'invalid_entries' => 0,
        'outliers_flagged' => 0
    ];

    foreach ($rawRows as $row) {
        // Ensure keys exist before checking
        $keys = array_keys($row);
        $firstCol = $keys[0] ?? null;
        
        // Skip empty core records
        if ($firstCol === null || empty(trim($row[$firstCol]))) {
            $metrics['missing_core_data']++;
            continue;
        }

        // Clean string fields (trim whitespace and normalize text)
        foreach ($row as $key => $value) {
            $row[$key] = trim($value);
            // Standardize capitalization for text description columns if needed
            if (stripos($key, 'name') !== false || stripos($key, 'station') !== false) {
                $row[$key] = ucwords(strtolower($row[$key]));
            }
        }

        // Validate numeric parameters like pH if present
        foreach ($row as $key => $value) {
            if (stripos($key, 'ph') !== false && !empty($value)) {
                if (is_numeric($value)) {
                    $pHVal = floatval($value);
                    // Standard environmental pH limits (0 to 14)
                    if ($pHVal < 0 || $pHVal > 14) {
                        $metrics['outliers_flagged']++;
                        $row[$key] = null; // Nullify physical impossibility
                    }
                } else {
                    $metrics['invalid_entries']++;
                    $row[$key] = null;
                }
            }
        }

        // Duplicate Detection using a composite signature of the row
        $rowSignature = implode('|', $row);
        if (isset($duplicateTracker[$rowSignature])) {
            $metrics['duplicates_removed']++;
            continue;
        }
        $duplicateTracker[$rowSignature] = true;

        $cleanedData[] = $row;
    }

    $finalCount = count($cleanedData);

    // 3. Export Cleaned Dataset
    $outputHandle = fopen($outputFile, 'w');
    if (!empty($cleanedData) && !empty($headers)) {
        fputcsv($outputHandle, $headers);
        foreach ($cleanedData as $row) {
            fputcsv($outputHandle, $row);
        }
    }
    fclose($outputHandle);

    // 4. Print Summary Report (For your documentation and tables)
    echo "   ----------------------------------------\n";
    echo "   📊 Results for this file:\n";
    echo "      - Initial Records     : {$initialCount}\n";
    echo "      - Final Cleaned Rows  : {$finalCount}\n";
    echo "      - Missing Skipped     : {$metrics['missing_core_data']}\n";
    echo "      - Duplicates Removed  : {$metrics['duplicates_removed']}\n";
    echo "      - Invalid Entries Fixed: {$metrics['invalid_entries']}\n";
    echo "      - Outliers Flagged    : {$metrics['outliers_flagged']}\n";
    echo "      - Saved As            : {$outputFile}\n";
    echo "   ----------------------------------------\n\n";
}

echo "=== ALL FILES PREPROCESSED SUCCESSFULLY ===\n";
?>