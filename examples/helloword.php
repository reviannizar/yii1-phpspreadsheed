<?php

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Hello World !');

$writer = new Xlsx($spreadsheet);
$outxls = 'hello world.xlsx';
//$writer->save($outxls);

header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
header('Pragma: public');
header('Cache-Control: max-age=0');
header('Content-Disposition: attachment; filename="'.$outxls.'"');
header('Content-type: application/vnd.ms-excel');
$writer->save('php://output');
