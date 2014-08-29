<?php
    require_once './XLStoCSVParser.php';
    $reader = new XLStoCSVParser();
    $reader->ParseXLStoCSV('PRAJS_OPT.xls', 'file1.csv');
?>