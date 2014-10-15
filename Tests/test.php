<?php

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
ini_set('memory_limit', '2048M');
ini_set('max_execution_time', 3600);

date_default_timezone_set("asia/shanghai");

require_once dirname(__FILE__) . '/../ExcelReader.php';

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

//$cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_in_memory_gzip;
//PHPExcel_Settings::setCacheStorageMethod($cacheMethod);

$columnDefines = array(
    array(
        'required' => true,
        'type' => 'int',
        'name' => 'databaseId',
        'key' => 'c1',
    ),

    array(
        'required' => true,
        'type' => 'string',
        'name' => 'useremail',
        'key' => 'c2',
    ),

    array(
        'required' => true,
        'type' => 'string',
        'name' => 'ebayListingID',
        'key' => 'c3',
    ),

    array(
        'required' => true,
        'type' => 'string',
        'name' => 'title',
        'key' => 'c4',
    ),

    array(
        'required' => true,
        'type' => 'date',
        'name' => 'startTime',
        'key' => 'c5',
    ),

    array(
        'required' => true,
        'type' => 'float',
        'name' => 'ConvertedStartPrice',
    ),
);

$columnMappings = array(
    "id" => "databaseId",
);

$testFiles = array('test_0.csv', 'test_1.csv');

$callStartTime = microtime(true);

foreach ($testFiles as $file) {
    echo "---------------------- Reading $file -------------------------" . EOL;
    $reader = new ExcelReader($file, $columnDefines, $columnMappings);
    $headers = $reader->getHeaders();

    echo "The headers of $file are: ";
    foreach ($headers as $header) {
        echo $header . '  ';
    }
    echo EOL;

    while (!$reader->finished()) {
        $data = $reader->read();

        if ($data['error']) {
            foreach ($data['error'] as $error) {
                echo "Error: $error" . EOL;
            }
        } else {
            foreach ($data['warn'] as $warn) {
                echo "Warn: $warn" . EOL;
            }

            echo 'Successfully read rows: ' . count($data['list']) . EOL;
            /*
            // echo the last time read data
            echo '<table>' . "\n";
            foreach ($data['list'] as $row) {
                echo '<tr>' . "\n";
                foreach ($row as $key => $value) {
                    echo '<td>' . "[$key]" . $value . '</td>' . "\n";
                }
                echo '</tr>' . "\n";
            }
            echo '</table>' . "\n";
            */
        }
    }

    $callEndTime = microtime(true);

    $callTime = $callEndTime - $callStartTime;

    echo EOL;
    echo "Cost time: ", sprintf('%.4f', $callTime), " seconds", EOL;
    echo 'Memory usage: ', (memory_get_usage(true) / 1024 / 1024), " MB", EOL;
    echo EOL . EOL . EOL;
}
