yii-phpspreadsheet
============

Menggunakan PhpSpreadsheet tanpa instalasi (seperti PHPExcel) bisa digunakan untuk native php yang disediakan contoh dalam folder examples dan extension untuk yii1 (Excel.php).


## Instalasi native php

1. Unzip content ke dalam folder root aplikasi
2. Untuk update phpspreadsheet download src [https://github.com/PHPOffice/PhpSpreadsheet] kemudian replace folder src dalam folder vendor/phpoffice/phpspreadsheet/src 

## Penggunaan

```php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Hello World !');

$writer = new Xlsx($spreadsheet);
$outxls = 'hello world.xlsx';
// Simpan dalam file
// $writer->save($outxls);

header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
header('Pragma: public');
header('Cache-Control: max-age=0');
header('Content-Disposition: attachment; filename="'.$outxls.'"');
header('Content-type: application/vnd.ms-excel');
$writer->save('php://output');
```


## Instalasi extension yii1

1. Unzip content ke dalam folder protected/extensions/phpexcel

2. Untuk update phpspreadsheet download src [https://github.com/PHPOffice/PhpSpreadsheet] kemudian replace folder src dalam folder vendor/phpoffice/phpspreadsheet/src 

<img src="img/folder.png" />

## Penggunaan

Extension untuk yii1 hanya menyediakan model menulis array secara langsung, cara ini memiliki kelebihan bisa export ke excel dalam jumlah ribuan row.

Bisa ditambahkan file excel sebagai template.

```php
	public function actionExcel()
	{
		$title='Test Excel';
		$tmpl='';
		$data=[
			['col A row 1','col B row 1','col C row 1',],
			['col A row 2','col B row 2','col C row 2',],
			['col A row 3','col B row 3','col C row 3',],
		];
		if($data){
			$this->widget('ext.phpexcel.Excel', ['data'=>$data,'title'=>"$title",'tempPath'=>file_exists($tmpl)?$tmpl:'']);
		}
	}

```
