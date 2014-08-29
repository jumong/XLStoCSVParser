<?php
require_once 'Excel/reader.php';

class XLStoCSVParser extends Spreadsheet_Excel_Reader
{
    var $_reader;
    
    function XLStoCSVParser() 
    {
        $this->_reader = new Spreadsheet_Excel_Reader();
        $this->_reader->setOutputEncoding('Windows-1251');
    }

    function ParseXLStoCSV($fileXLS, $fileCSV)
    {
        $this->_reader->read($fileXLS);
        $this->_fp = fopen($fileCSV, 'w');
        foreach($this->_reader->sheets as $data)
        {
            if(isset($data['cells']))
            {
                foreach($data['cells'] as $row) 
                {
                   foreach($row as $cell)
                   {
                       $arrayToCSV = new SplFixedArray(19);
                       if($this->slashCount($cell) == 0 && array_search($cell, $row) != 3){
                           break;
                       }
                       if(array_search($cell, $row) == 3){
                           $seasonvariable = ($this->slashCount($cell) == 1 ? $this->getStringAfterCh($cell, '/') : null );
                           continue;
                       }
                       $arrayToCSV[7] = $seasonvariable;
                       if($this->slashCount($cell) != 0){
                           $arrayFromFirstCell = $this->getFirstCell($cell);
                           $arrayToCSV[4] = (isset($arrayFromFirstCell[0])? $arrayFromFirstCell[0] : null);
                           $arrayToCSV[5] = (isset($arrayFromFirstCell[1])? $arrayFromFirstCell[1] : null);
                           $arrayToCSV[6] = (isset($arrayFromFirstCell[2])? $arrayFromFirstCell[2] : null);
                       }
                       $arrayToCSV[2] = $row[2];
                       $arrayToCSV[3] = $row[3];
                       $arrayToCSV[12] = $row[9];
                       $arrayToCSV[13] = $row[8];
                       $arrayToCSV[15] = $row[4];
                       fputcsv($this->_fp, $arrayToCSV->toArray(), ';', ' ');
                   }
               }
            }
        }
        fclose($this->_fp);
    }
    
    function slashCount($str) //returns number of '/' in string
    {
        $flag = 0;
        for($i=0; $i<strlen($str); $i++){
            if($str[$i]=='/'){
                $flag ++;
            }
        }
        return $flag;
    }

    function getStringAfterCh($str, $ch) //get string After / - R10/simf -> 'simf'
    {
        for($i=0; $i<strlen($str); $i++){
            if($str[$i] == $ch && $i != strlen($str)){
                return substr($str, $i+1);
            }
        }
    }

    function getFirstCell($str) //returns array of sizes
    {
        $array = array();
        $str1 = '';
        for($i=0; $i<strlen($str); $i++){
            if($str[$i]=='.'){
                $str[$i] = ',';
            }
            if($str[$i] == '/' || $i == strlen($str)-1){
                if($i == strlen($str)-1 && $str[$i] != '/'){
                    $str1 .= $str[$i]; 
                }
                array_push($array, $str1);
                $str1 = '';
                continue;
            }
            $str1 .= $str[$i]; 
        }
        return $array;
    }
}
?>
