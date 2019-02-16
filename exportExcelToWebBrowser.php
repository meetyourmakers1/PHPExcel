<?php

$dir = dirname(__FILE__);
require $dir.'/dbconfig.php';
require $dir.'/db.php';
require $dir.'/PHPExcel/PHPExcel.php';

$db = new db($dbconf);
$excel = new PHPExcel();
for($i = 1;$i <= 3;$i++)
{
    if($i > 0)
    {
        $excel->createSheet();
    }
    $excel->setActiveSheetIndex($i);
    $sheet = $excel -> getActiveSheet();
    $data = $db->getDataByGrade($i);
    $title = $i.'年级';
    $sheet->setTitle($title);
    $sheet->setCellValue('A1','姓名')
          ->setCellValue('B1','年级')
          ->setCellValue('C1','班级')
          ->setCellValue('D1','成绩');
    $j = 2;
    foreach($data as $k => $v)
    {
        $sheet->setCellValue('A'.$j,$v['name'])
              ->setCellValue('B'.$j,$v['grade'])
              ->setCellValue('C'.$j,$v['class'])
              ->setCellValue('D'.$j,$v['score']);
        $j++;
    }
}
$phpExcel = PHPExcel_IOFactory::createWriter($excel,'Excel5');
$phpExcel->save($dir.'/export.xls');

//exportExcelToWebBrowser("Excel5",'webBrowser.excel05.xls');
//$phpExcel->save("php://output");

function exportExcelToWebBrowser($excelType,$excelName)
{
    if($excelType == 'Excel5')
    {
        header('Content-Type: application/vnd.ms-excel');
    }
    else
    {
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    }
    header('Content-Disposition: attachment;filename="'.$excelName.'"');
    header('Cache-Control: max-age=0');
}


