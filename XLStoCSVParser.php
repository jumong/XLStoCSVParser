<?php
require_once 'Excel/reader.php';

class XLStoCSVParser extends Spreadsheet_Excel_Reader
{
    var $_reader;
    
    function XLStoCSVParser() 
    {
        $this->_reader = new Spreadsheet_Excel_Reader();
        $this->_reader->setOutputEncoding('UTF-8');
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
                       if($this->slashCount($cell) == 0 && array_search($cell, $row) != 3){
                           break;
                       }
                       
                       $arrayToCSV = new SplFixedArray(19);
                       if(array_search($cell, $row) == 3 && sizeof($row) == 1){
                            if(strtolower($cell)== 'грузовые шины'){
                                $tiresvariable = 'грузовой';
                                $seasonvariable = null;
                            }
                            elseif (strtolower($cell)== 'легкогрузовые шины'){
                                $tiresvariable = 'легковой';
                                $seasonvariable = null;
                            }
                            else{
                                $tiresvariable = null;
                            }
                            
                            $seasonvariable = $this->formatSeasons(($cell[0]=='R' ? $this->getStringAfterCh($cell, array('/', ' ')) : null ));
                            continue;
                       }
                       $arrayToCSV[7] = $seasonvariable;
                       $arrayToCSV[8] = $tiresvariable;
                       if($this->slashCount($cell) != 0){
                           $arrayFromFirstCell = $this->getFirstCell($cell);
                           $arrayToCSV[4] = (isset($arrayFromFirstCell[0])? $arrayFromFirstCell[0] : null);
                           $arrayToCSV[5] = (isset($arrayFromFirstCell[1])? $arrayFromFirstCell[1] : null);
                           $arrayToCSV[6] = (isset($arrayFromFirstCell[2])? $arrayFromFirstCell[2] : null);
                       }
                       $arrayToCSV[2] = isset($row[2]) ? $row[2] : null;
                       $arrayToCSV[3] = isset($row[3]) ? $row[3] : null;
                       $arrayToCSV[12] = isset($row[9]) ? $row[9] : null;
                       $arrayToCSV[13] = isset($row[8]) ? $row[8] : null;
                       $arrayToCSV[15] = isset($row[4]) ? $row[4] : null;
                       date_default_timezone_set('Europe/Moscow');
                       $arrayToCSV[16] = date('d/m/Y G:i:s', time());
                       $arrayToCSV[18] = $arrayToCSV[2].' '.$arrayToCSV[3].' '.$arrayToCSV[4].'/'.$arrayToCSV[5].' '.($arrayToCSV[6] != null? 'R'.$arrayToCSV[6].' ': '').$arrayToCSV[10].$arrayToCSV[9];
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
    
    function formatSeasons($str)
    {
        $strlow = strtolower($str);
        switch ($strlow) {
            case "лето-всесезонка":
              return "летняя-всесезонная";
              break;
            case "зима":
              return "зимняя";
              break;
            default:
              return $str;
              break;
        }
    }

    function getStringAfterCh($str, $arr) //get string After / - R10/simf -> 'simf'
    {
        for($i=0; $i<strlen($str); $i++){
            if(in_array($str[$i],$arr) && $i != strlen($str)){
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
