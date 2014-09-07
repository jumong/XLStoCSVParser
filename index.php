<?php
    require_once 'XLStoCSVParser.php';
    $reader = new XLStoCSVParser();
    $reader->ParseXLStoCSV('price_04-09-2014.xls', 'price_04-09-2014.csv');
?>