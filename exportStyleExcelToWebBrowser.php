<?php

$dir = dirname(__FILE__);
require $dir.'/dbconfig.php';
require $dir.'/db.php';
require $dir.'/PHPExcel/PHPExcel.php';

$db = new db($dbconf);
$excel = new PHPExcel();


$sheet = $excel->getActiveSheet();

$sheet->getDefaultStyle()
      ->getAlignment()
      ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
      ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$sheet->getDefaultStyle()
      ->getFont()
      ->setName('微软雅黑')
      ->setSize('16');
$sheet->getStyle('A2:Z2')
      ->getFont()
      ->setSize(20)
      ->setBold(True);
$sheet->getStyle('A3:Z3')
      ->getFont()
      ->setSize(18)
      ->setBold(True);

$grade = $db->getAllGrade();
$index = 0;
foreach($grade as $g_k => $g_v)
{
    $gradeIndex = getColumn($index*2);

    $sheet->setCellValue($gradeIndex.'2','高'.$g_v['grade']);

    $class = $db->getAllClass($g_v['grade']);

    foreach($class as $c_k => $c_v)
    {
        $nameIndex = getColumn($index*2);
        $scoreIndex = getColumn($index*2+1);

        $sheet->setCellValue($nameIndex.'3',$c_v['class'].'班');

        $sheet->mergeCells($nameIndex.'3:'.$scoreIndex.'3');

        $sheet->getStyle($nameIndex.'3:'.$scoreIndex.'3')
              ->getFill()
              ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
              ->getStartColor()
              ->setRGB('FFFF00');

        $borderStyle = getBorder('000');
        $sheet->getStyle($nameIndex.'3:'.$scoreIndex.'3')
              ->applyFromArray($borderStyle);

        $sheet->getStyle($nameIndex.'4')
              ->getAlignment()
              ->setWrapText(true);

        $sheet->setCellValue($nameIndex.'4',"姓\n名")
              ->setCellValue($scoreIndex.'4','成绩');

        $sheet->getStyle($nameIndex.'4:'.$scoreIndex.'4')
              ->getFill()
              ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
              ->getStartColor()
              ->setRGB('00FF00');

        $sheet->getStyle($nameIndex.'4')
              ->applyFromArray($borderStyle);
        $sheet->getStyle($scoreIndex.'4')
              ->applyFromArray($borderStyle);

        $student = $db->getAllStudent($g_v['grade'],$c_v['class']);
        $sheet->getStyle($scoreIndex)
              ->getNumberFormat()
              ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_TEXT);
        $j = 5;
        foreach($student as $k => $v)
        {
            /*$sheet->setCellValue($nameIndex.$j,$v['name'])
                  ->setCellValue($scoreIndex.$j,$v['score']);*/
            $sheet->setCellValue($nameIndex.$j,$v['name'])
                  ->setCellValueExplicit($scoreIndex.$j,'01234567891011121314151617181920',PHPExcel_Cell_DataType::TYPE_STRING);
            $j++;
        }
        $index++;
    }
    $endGradeIndex = getColumn($index*2-1);

    $sheet->mergeCells($gradeIndex.'2:'.$endGradeIndex.'2');

    $sheet->getStyle($gradeIndex.'2:'.$endGradeIndex.'2')
          ->getFill()
          ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
          ->getStartColor()
          ->setRGB('FF0000');

    $sheet->getStyle($gradeIndex.'2:'.$endGradeIndex.'2')
          ->applyFromArray($borderStyle);
}

function getColumn($index)
{
    $array = range('A','Z');
    return $array[$index];
}

function getBorder($color)
{
    $styleArray = array(
        'borders' => array(
            'outline' => array(
                'style' => PHPExcel_Style_Border::BORDER_THICK,
                'color' => array('rgb' => $color)
            )
        )
    );
    return $styleArray;
}


$phpExcel = PHPExcel_IOFactory::createWriter($excel,'Excel5');
//$phpExcel->save($dir.'/export.xls');

exportExcelToWebBrowser("Excel5",'webBrowser.excel05.xls');
$phpExcel->save("php://output");

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