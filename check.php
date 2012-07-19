<?php
error_reporting(E_ALL);
ini_set('max_execution_time', 3000);
date_default_timezone_set('Europe/London');

/** Include PHPExcel_IOFactory */
require_once '../Classes/PHPExcel/IOFactory.php';

$objPHPExcel = new PHPExcel();


if (!file_exists("data.xls")) {
	exit("data.xls not found. Please load it in the directory.<br/>" . PHP_EOL);
}

if (!file_exists("master.xls")) {
	exit("master.xls not found. Please load it in the directory.<br/>" . PHP_EOL);
}

//echo date('H:i:s') , " Loading Master Database<br/>" , PHP_EOL;
$objPHPExcelMaster = PHPExcel_IOFactory::load("master.xls");

//echo date('H:i:s') , " Loading Data Database<br/>" , PHP_EOL;
$objPHPExcelData = PHPExcel_IOFactory::load("data.xls");

$BrandDataBase = PHPExcel_IOFactory::load("brands.xls");

//Preparing the Array of Data from Master Database
$rowIterator = $objPHPExcelMaster->getActiveSheet()->getRowIterator();

$sheet1 = $objPHPExcelMaster->getActiveSheet();
$arrayMaster = array();
foreach($rowIterator as $row){
    $rowIndex = $row->getRowIndex ();
    $cell = $sheet1->getCell('A' . $rowIndex);
    $arrayMaster[$rowIndex]['A']= $cell->getCalculatedValue();
   
}
//print_r($arrayMaster);


//Preparing the Array of Data from data Database
$rowIterator = $objPHPExcelData->getActiveSheet()->getRowIterator();

$sheet2 = $objPHPExcelData->getActiveSheet();
$arrayData = array();
foreach($rowIterator as $row){
    $rowIndex = $row->getRowIndex ();
    $cell = $sheet2->getCell('A' . $rowIndex);
    $arrayData[$rowIndex]['A']= $cell->getCalculatedValue();
    
    
}
//print_r($arrayData);

//Preparing the Array of Data from Brands Database
$sheet3 = $BrandDataBase->getActiveSheet();
$rowIterator = $BrandDataBase->getActiveSheet()->getRowIterator();
$brands = array();
foreach($rowIterator as $row){
    $rowIndex = $row->getRowIndex ();
	$cell = $sheet3->getCell('A'.$rowIndex);
	$brands[$rowIndex]=$cell->getCalculatedValue();
}
$models = array();
foreach($rowIterator as $row){
    $rowIndex = $row->getRowIndex ();
	$cell = $sheet3->getCell('B'.$rowIndex);
	$models[$rowIndex]=$cell->getCalculatedValue();
}




$rowsData = count($arrayData);
$rowsMaster = count($arrayMaster);
$rowsBrands = count($brands);
$rowsModels = count($models);

//$brand = array('Samsung', 'Micromax', 'Nokia');

for($i=1; $i<=$rowsData; $i++)
//for($i=1; $i<=10; $i++)
  {
   
   $finalMatch=0;
   for($j=1; $j<=$rowsMaster; $j++)
   //for($j=1; $j<=10; $j++)
   {     
    $match=0; 
	$DataBrand="";
	$DataModel="";
	$DataNumber=""; 
    	
	$DataStr= $arrayData[$i]['A'];
    echo "<br/><<".$i.">> DataStr : ".$DataStr."<br/>";
    
     foreach ($brands as $DB) 
	  {
       if (preg_match("/$DB/i", $DataStr))
	   {
        $DataBrand=$DB;
        break;
       }
      }
     echo "Data Brand= ".$DataBrand."<br/>";
	 
     $str1= str_ireplace($DataBrand, '', $DataStr);
     //echo "Str1 : ".$str1."<br/>";

     foreach ($models as $DM) 
	 {
      if (preg_match("/$DM/i", $str1))
	   {
        $DataModel=$DM;
        break;
       }
     }

     echo "Data Model= ".$DataModel."<BR/>";
	
     $DataNumber= str_ireplace($DataModel, '', $str1);
     echo "Data Model Number= ".$DataNumber."<BR/>";
	 
	 
	 
	 
	 
	 
	$MasterStr= $arrayMaster[$j]['A'];
        
     foreach ($brands as $MB) 
	  {
       if (preg_match("/$MB/i", $MasterStr))
	   {
        $MasterBrand=$MB;
        break;
       }
      }
     echo "Master Brand= ".$MasterBrand."<br/>";
	 
     $str1= str_ireplace($MasterBrand, '', $MasterStr);
     //echo $str1."<br/>";

     foreach ($models as $MM) 
	 {
      if (preg_match("/$MM/i", $str1))
	   {
        $MasterModel=$MM;
        break;
       }
     }

     echo "Master Model= ".$MasterModel."<BR/>";
	
     $MasterNumber = str_ireplace($MasterModel, '', $str1);
     echo "Master Model Number= ".$MasterNumber."<BR/>";
	 echo "-------------------------------------<br/>";

	
        if($DataBrand==$MasterBrand)
        $match=$match+20; echo $match;
        if($DataModel==$MasterModel)
        $match=$match+20; echo $match;
        if($DataNumber==$MasterNumber)
        $match=$match+20; echo $match;
        if($DataBrand==$MasterBrand&&$DataModel==$MasterModel)
        $match=$match+10; echo $match;
        if($DataBrand==$MasterBrand&&$DataNumber==$MasterNumber)
        $match=$match+10; echo $match;
        if($DataNumber==$MasterNumber&&$DataModel==$MasterModel)
        $match=$match+10; echo $match;
        if($DataBrand==$MasterBrand&&$DataModel==$MasterModel&&$DataNumber==$MasterNumber)
        $match=$match+10;  echo $match;
		echo "<br/>-------------------------------------<br/>";    
        if($match>$finalMatch)
        $finalMatch=$match;
    }
	
   echo "=====================================";  
   echo "<br/>(".$i.")->".$arrayData[$i]['A']." - match % = ".$finalMatch."<br/>"; 
   echo "=====================================<br/>";  
   
	
	$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A'.$i, $arrayData[$i]['A'])
            ->setCellValue('B'.$i, 'match % ='.$finalMatch)
			->setCellValue('C'.$i, $DataBrand)
			->setCellValue('D'.$i, $DataModel)
			->setCellValue('E'.$i, $DataNumber);
	if($finalMatch==100)
	{
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$i,"Fully Matched");
	}
	elseif($finalMatch==50)
	{
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$i,"Half Matched");
	}
	else
	{
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$i,"Partially Matched");
	}
      
  } 
  
  //$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
  //$objWriter->save("Sample-Output-New.xls");

?>