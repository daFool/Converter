#!/bin/php 
<?php

if ($argc != 3) {
    die("Usage: {$argv[0]} <excel> <csv> {$argc}".PHP_EOL);
}

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;

if(!preg_match('/^.*\.(.{3,4})$/', $argv[1], $m)) {
    die("Can't determine file type".PHP_EOL);
}

if(!preg_match('/^(.*)\.(.{3,4})$/', $argv[2], $n)) {
    die("Uh, can't parse target name, it should be like name.csv".PHP_EOL);
}
$body = $n[1];
$suffix = $n[2];

switch($m[1]) {
    case "xlsx":
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        break;
    case "xls":
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
        break;
    case "ods":
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Ods();
        break;
    default:
        die("Probably unsupported type {$m[1]}".PHP_EOL);
        break;
}
$spreadsheet = $reader->load($argv[1]);

$sheets=$spreadsheet->getSheetCount();
$sheetNames = $spreadsheet->getSheetNames();
foreach ($sheetNames as $snro=>$sn) {
    $filename = ($snro==0 && $sheets ==1) ? sprintf("%s.%s", $body, $suffix) : sprintf("%s_%s.%s", $body, $sn, $suffix);
    if (file_exists($filename)) {
        die("Target file {$filename} exists!".PHP_EOL);
    }
    $fp = fopen($filename, "w");
    if($fp===false) {
        die("Unable to open file {$filename}");
    }
    $sheet=$spreadsheet->getSheetByName($sn)->toArray();    
    printf("Generating file: %s.%s", $filename, PHP_EOL);
    foreach($sheet as $i=>$row) {
        $s="";
        foreach($row as $j=>$column) {
            if($j===0) {
                $s.=$column;
            } else {
                $s.=";".$column;
            }
        }
        fprintf($fp, "%s%s", $s, PHP_EOL);
  //      printf("%d: %s%s", $i, $s, PHP_EOL);
    }
    fclose($fp);
    unset($sheet);
}