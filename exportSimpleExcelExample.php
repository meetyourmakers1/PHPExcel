<?php

require dirname(__FILE__).'/PHPExcel/PHPExcel.php';

$excel = new PHPExcel();  //新建Excel文件

$sheet = $excel->getActiveSheet();  //新建sheet
$sheet->setTitle('excel1');  //sheet重命名

$sheet->setCellValue('A1','id')
      ->setCellValue('B1','姓名');  //填充数据
$sheet->setCellValue('A2','1')
      ->setCellValue('B2','张三');  //填充数据

$phpExcel = PHPExcel_IOFactory::createWriter($excel,'Excel2007');  //生成Excel
$phpExcel->save(dirname(__FILE__).'/excel1.xlsx');  //保存Excel