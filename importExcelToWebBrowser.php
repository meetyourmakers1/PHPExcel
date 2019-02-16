<?php

$dir = dirname(__FILE__);
require $dir."/PHPExcel/PHPExcel/IOFactory.php";
$filename = $dir."/excel1.xlsx";
$fileType=PHPExcel_IOFactory::identify($filename);
$objReader=PHPExcel_IOFactory::createReader($fileType);
$sheetName=array("excel1");
$objReader->setLoadSheetsOnly($sheetName);
$objExcel=$objReader->load($filename);
foreach($objExcel->getWorksheetIterator() as $sheet){
    foreach($sheet->getRowIterator() as $row){
        /*if($row->getRowIndex()<2){
            continue;
        }*/
        foreach($row->getCellIterator() as $cell){
            $data=$cell->getValue();
            echo $data." ";
        }
        echo '<br/>';
    }
    echo '<br/>';
}