<?php

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
ini_set('memory_limit', '2048M');
ini_set('max_execution_time', 3600);

require_once dirname(__FILE__) . '/../ExcelReader.php';

//$cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_in_memory_gzip;
//PHPExcel_Settings::setCacheStorageMethod($cacheMethod);

$callStartTime = microtime(true);

$reader = new ExcelReader('test.xls', null);
while (!$reader->finished()) {
    $data = $reader->read();
    // ...
}

$callEndTime = microtime(true);

$callTime = $callEndTime - $callStartTime;

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

echo 'Call time to read Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
// Echo memory usage
echo 'Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;

echo 'Total rows: ' . count($data);


// echo the last time read data

echo '<table>' . "\n";
foreach ($data as $row) {
    echo '<tr>' . "\n";
    foreach ($row as $cell) {
        echo '<td>' . $cell . '</td>' . "\n";
    }
    echo '</tr>' . "\n";
}
echo '</table>' . "\n";