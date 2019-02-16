<?php

$dir = dirname(__FILE__);
require $dir.'/PHPExcel/PHPExcel.php';


$excel = new PHPExcel();


$sheet = $excel->getActiveSheet();
$sheetDrawing = new PHPExcel_WorkSheet_Drawing();
$sheetDrawing->setPath($dir."/image/image.jpg");
$sheetDrawing->setCoordinates("C3");
$sheetDrawing->setWidth(500);
$sheetDrawing->setHeight(500);
$sheetDrawing->setOffsetX(15)
             ->setOffsetY(10);
$sheetDrawing->setWorksheet($sheet);

$sheetText=new PHPExcel_RichText();
$sheetText->createText("慕课网(https://www.imooc.com/)");
$sheetFont=$sheetText->createTextRun("专注做好IT技能教育的MOOC，符合互联网发展潮流接地气儿的MOOC。我们免费，我们只教有用的，我们专心做教育。");
$sheetFont->getFont()
          ->setSize(12)
          ->setBold(True)
          ->setColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_RED));
$sheetText->createText('让更多热爱互联网的同学来慕课网学习，多年以后，圈子里一批技术牛说：我在慕课网学习过，这就够了。');
$sheet->getCell('A2')
      ->setValue($sheetText);

$sheet->mergeCells('A2:Z2');
$sheet->getComment('A2')
      ->getText()
      ->createTextRun('Van:\r\n慕课网\n\nimooc');

$sheet->setCellValue("J1","慕课网");
$sheet->getStyle("J1")
      ->getFont()
      ->setUnderline(true)
      ->setColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_BLUE));
$sheet->getCell("J1")
      ->getHyperlink()
      ->setUrl("http://www.imooc.com");


$phpExcel = PHPExcel_IOFactory::createWriter($excel,'Excel2007');
//$phpExcel->save($dir.'/export.xls');

exportExcelToWebBrowser("Excel2007",'webBrowser.excel2007.xlsx');
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