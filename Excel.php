<?php
/**
 * EExcel class file.
 * @author Abu dzunnuraini <almprokdr@gmail.com>
 * @copyright Copyright &copy; Abu dzunnuraini 2020
 * Tidak bisa export pdf
 * <pre>
 * $this->widget('ext.phpexcel.Excel', array(
 *     'filename'=>'';
 *     'outType'=>'xlsx' // default
 *     'temppath'=>'view/report/xx.xls',
 *     'data'=>array(),
 * ));
 * </pre>
 * @data
 */
 
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\IOFactory;
 
define('_SET_MEMORY_LIMIT_','1024M');

class Excel extends CWidget{
	
	public $creator		= 'Abu dzunnuraini';
	public $title		= '';
	public $subject		= '';
	public $description	= '';
	public $category	= '';
	public $keywords	= '';

	public $tmpType		= ['Xls','Xlsx'];
	public $outType		= 'xlsx';
	
	public $data		= [];
	public $extdata		= [];
	public $startRow	= 1;
	public $tempPath	= '';
	public $filename	= null; //export FileName //false;
	public $stream		= true; //stream to browser
	
	public $memLimit	= false;
	public $spreadsheet	= null;

	public function init(){
		$this->title=($this->title!='')?$this->title:Yii::app()->getController()->getPageTitle();
		$this->spreadsheet = new PhpSpreadsheet\Spreadsheet();
		$this->spreadsheet->getProperties()
			->setCreator($this->creator)
			->setTitle($this->title)
			->setSubject($this->subject)
			->setDescription($this->description)
			->setKeywords($this->keywords)
			->setCategory($this->category);
			
		if($this->tempPath!=''){
			$ext=pathinfo($this->tempPath); $ext=strtolower($ext['extension'])=='xls'?0:1; 
			$reader = IOFactory::createReader($this->tmpType[$ext]);
			$this->spreadsheet = $reader->load($this->tempPath);
		}
		if($this->memLimit) ini_set('memory_limit',_SET_MEMORY_LIMIT_);
	}

	public function run(){
		if(!empty($this->data)){ 
			$worksheet = $this->spreadsheet->getActiveSheet();
			$worksheet->fromArray($this->data, null, 'A'.$this->startRow);
		}
		if(!empty($this->extdata)) $this->SetExtData($this->extdata, $this->startRow);
		$this->createOutput();
	}
	
	private function SetExtData($data, $startRow){
		$index=0;
		foreach($data as $row){
			extract($row);
			$worksheet = $this->spreadsheet->getSheet($index);
			$worksheet->setTitle($name,false)->fromArray($data, null, 'A'.$startRow);		
		}
	}

	public function createOutput(){
		$writer = $this->outType=='xlsx' ? new Xlsx($this->spreadsheet): new Xls($this->spreadsheet);
		if(!$this->stream){ 
			$writer->save($this->filename); 
		}else {
			$this->filename = $this->filename==null ? $this->title : $this->filename;	
			$this->cleanOutput();
			header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
			header('Pragma: public');
			header('Cache-Control: max-age=0');
			header('Content-Disposition: attachment; filename="'.$this->filename.'.'.$this->outType.'"');
			header('Content-type: application/vnd.ms-excel');
			$writer->save('php://output');
			Yii::app()->end();
		}
	}

	private static function cleanOutput(){
		for($level=ob_get_level();$level>0;--$level){@ob_end_clean();}
	}

}
