<?php
namespace Home\Action;
class InsertAction extends BaseAction{
	public function index(){
		$this->display('default/excel');
	}
	public function ddd(){
		if (! empty ( $_FILES ['file_stu'] ['name'] )){
		    $tmp_file = $_FILES ['file_stu'] ['tmp_name'];
		    $file_types = explode ( ".", $_FILES ['file_stu'] ['name'] );
		    $file_type = $file_types [count ( $file_types ) - 1];
		     /*判别是不是.xls文件，判别是不是excel文件*/
		    if (strtolower ( $file_type ) != "xls"){
		        $this->error ( '不是Excel文件，重新上传' );
		    }
		    /*设置上传路径*/
		    $savePath = './Upload/Excel/';
		    /*以时间来命名上传的文件*/
		    $str = date ( 'Ymdhis' ); 
		    $file_name = $str . "." . $file_type;
		     /*是否上传成功*/   
		    if (!copy($tmp_file,$savePath . $file_name )) {
		        $this->error ( '上传失败' );
		    }
		    else{
		    	$result = $this->importExecl($savePath.$file_name);
		    	$dataa=$result['data'][0]['Content'];
		    	foreach ($dataa as $key => $value) {
		    		$m=M('carserver.users');
		    		$data['loginName']='jh'.$value['0'].'_a';
		    		$data['loginSecret']=9999;
		    		$data['loginPwd']='7e3bb63f6e447bf1c9e4737c2d6a1a8b';
		    		$data['userName']=$value['1'];
		    		$data['userPhone']=$value['0'];
		    		$data['cid']=102;
		    		$row=$m->add($data);

		    		$m2=M('yuexue.users');
		    		$data2['loginName']='jh'.$value['0'].'_a';
		    		$data2['loginSecret']=9999;
		    		$data2['loginPwd']='7e3bb63f6e447bf1c9e4737c2d6a1a8b';
		    		$data2['userName']=$value['1'];
		    		$data2['userPhone']=$value['0'];
		    		$data2['cid']=102;
		    		$data2['yxprice']=999;
		    		$row2=$m2->add($data2);

		    		$m3=M('company.gx');
		    		$data3['yuserId']=$row2;
		    		$data3['cuserId']=$row;
		    		$data3['comId']=102;
		    		$data3['userName']=$value['1'];
		    		$data3['sn']=$value['0'];
		    		$row3=$m3->add($data3);
		    	}
		    }
		}
	}
	 public function importExecl($file){ 
        if(!file_exists($file)){ 
            return array("error"=>0,'message'=>'file not found!');
        } 
        Vendor("PHPExcel.PHPExcel.IOFactory"); 
        $objReader =\PHPExcel_IOFactory::createReader('Excel5'); 
        try{
            $PHPReader = $objReader->load($file);
        }catch(Exception $e){}
        if(!isset($PHPReader)) return array("error"=>0,'message'=>'read error!');
        $allWorksheets = $PHPReader->getAllSheets();
        $i = 0;
        foreach($allWorksheets as $objWorksheet){
            $sheetname=$objWorksheet->getTitle();
            $allRow = $objWorksheet->getHighestRow();//how many rows
            $highestColumn = $objWorksheet->getHighestColumn();//how many columns
            $allColumn =\PHPExcel_Cell::columnIndexFromString($highestColumn);
            $array[$i]["Title"] = $sheetname; 
            $array[$i]["Cols"] = $allColumn; 
            $array[$i]["Rows"] = $allRow; 
            $arr = array();
            $isMergeCell = array();
            foreach ($objWorksheet->getMergeCells() as $cells) {//merge cells
                foreach (\PHPExcel_Cell::extractAllCellReferencesInRange($cells) as $cellReference) {
                    $isMergeCell[$cellReference] = true;
                }
            }
            for($currentRow = 1 ;$currentRow<=$allRow;$currentRow++){ 
                $row = array(); 
                for($currentColumn=0;$currentColumn<$allColumn;$currentColumn++){;                
                    $cell =$objWorksheet->getCellByColumnAndRow($currentColumn, $currentRow);
                    $afCol = \PHPExcel_Cell::stringFromColumnIndex($currentColumn+1);
                    $bfCol = \PHPExcel_Cell::stringFromColumnIndex($currentColumn-1);
                    $col = \PHPExcel_Cell::stringFromColumnIndex($currentColumn);
                    $address = $col.$currentRow;
                    $value = $objWorksheet->getCell($address)->getValue();
                    if(substr($value,0,1)=='='){
                        return array("error"=>0,'message'=>'can not use the formula!');
                        exit;
                    }
                    if($cell->getDataType()==\PHPExcel_Cell_DataType::TYPE_NUMERIC){
                        $cellstyleformat=$cell->getParent()->getStyle( $cell->getCoordinate() )->getNumberFormat();
                        $formatcode=$cellstyleformat->getFormatCode();
                        if (preg_match('/^([$[A-Z]*-[0-9A-F]*])*[hmsdy]/i', $formatcode)) {
                            $value=gmdate("Y-m-d", \PHPExcel_Shared_Date::ExcelToPHP($value));
                        }else{
                            $value=\PHPExcel_Style_NumberFormat::toFormattedString($value,$formatcode);
                        }                
                    }
                    if($isMergeCell[$col.$currentRow]&&$isMergeCell[$afCol.$currentRow]&&!empty($value)){
                        $temp = $value;
                    }elseif($isMergeCell[$col.$currentRow]&&$isMergeCell[$col.($currentRow-1)]&&empty($value)){
                        $value=$arr[$currentRow-1][$currentColumn];
                    }elseif($isMergeCell[$col.$currentRow]&&$isMergeCell[$bfCol.$currentRow]&&empty($value)){
                        $value=$temp;
                    }
                    $row[$currentColumn] = $value; 
                } 
                $arr[$currentRow] = $row; 
            } 
            $array[$i]["Content"] = $arr; 
            $i++;
        } 
        // spl_autoload_register(array('Think','autoload'));//must, resolve ThinkPHP and PHPExcel conflicts
        unset($objWorksheet); 
        unset($PHPReader); 
        unset($PHPExcel); 
        unlink($file); 
        return array("error"=>1,"data"=>$array); 
    }
};
?>