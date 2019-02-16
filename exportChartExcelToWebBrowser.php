<?php

$dir = dirname(__FILE__);
require $dir.'/PHPExcel/PHPExcel.php';


$excel = new PHPExcel();


$sheet = $excel->getActiveSheet();
$array = array(
    array('','一班','二班','三班','四班','五班','六班','七班','八班','九班','十班','十一班','十二班','十三班','十四班','十五班','十六班','十七班'),
    array('不及格',43,12,34,65,45,87,23,12,93,50,65,86,87,23,12,93,50),
    array('良好',21,39,65,23,84,45,82,84,35,83,31,86,23,64,72,56,23),
    array('优秀',65,31,86,31,14,23,64,72,56,23,21,39,65,23,87,45,82,)
);
$sheet->fromArray($array);
$labels = array(
    new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$A$2',null,1),
    new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$A$3',null,1),
    new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$A$4',null,1)
    );
$xLabels = array(
    new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$B$1:$R$1',null,10)
);
$datas = array(
    new PHPExcel_Chart_DataSeriesValues('Number','Worksheet!$B$2:$R$2',null,10),
    new PHPExcel_Chart_DataSeriesValues('Number','Worksheet!$B$3:$R$3',null,10),
    new PHPExcel_Chart_DataSeriesValues('Number','Worksheet!$B$4:$R$4',null,10)
);
$series = array(
    new PHPExcel_Chart_DataSeries(
        PHPExcel_Chart_DataSeries::TYPE_LINECHART,
        PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
        range(0,count($labels)-1),
        $labels,
        $xLabels,
        $datas
    )
);
$title=new PHPExcel_Chart_Title("高一学生成绩");
$yAxisLabel=new PHPExcel_Chart_Title("人数");
$legend=new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT,null,false);
$layout=new PHPExcel_Chart_Layout();
$layout->setShowVal(true);
$plotAreas=new PHPExcel_Chart_PlotArea($layout,$series);

$chart=new PHPExcel_Chart(
    'line_chart',
    $title,
    $legend,
    $plotAreas,
    true,
    false,
    null,
    $yAxisLabel
);
$chart->setTopLeftPosition("A7")->setBottomRightPosition("S25");
$sheet->addChart($chart);

$phpExcel = PHPExcel_IOFactory::createWriter($excel,'Excel2007');
$phpExcel->setIncludeCharts(true);
$phpExcel->save($dir.'/export.xls');


exportExcelToWebBrowser("Excel2007",'webBrowser.chart2007.xlsx');
$phpExcel->save("php://output");

function exportExcelToWebBrowser($excelType,$excelName)
{
    if ($excelType == 'Excel5') {
        header('Content-Type: application/vnd.ms-excel');
    } else {
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    }
    header('Content-Disposition: attachment;filename="' . $excelName . '"');
    header('Cache-Control: max-age=0');
}