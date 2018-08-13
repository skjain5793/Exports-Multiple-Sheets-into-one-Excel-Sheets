<?php
include("dbinfo.inc.php");
require_once 'PHPExcel-1.8/Classes/PHPExcel.php';
require_once 'PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';
//require_once("codelibrary/inc/variables.php");


$temp_year_start = $_REQUEST['temp_year_start'];
$temp_year_end = $_REQUEST['temp_year_end'];   

$company = $_REQUEST['company'];
$state = $_REQUEST['state'];
$region = $_REQUEST['region'];

// Confirmed tab

$query="SELECT r.*, DATE_FORMAT(r.cycle_start_date, '%m/%d/%Y') AS start_date, u.firstname,u.lastname, u.state, u.company_id, region_states.region, ur.firstname rfi_firstname,ur.lastname rfi_lastname  

FROM `rater_1qa_tracking` r 

LEFT OUTER JOIN users u ON r.rater_id = u.id  

LEFT OUTER JOIN users ur ON r.rfi_id = ur.id  

LEFT OUTER JOIN region_states ON u.`state` = region_states.state_short 

WHERE u.raterstatus NOT IN ('Terminated','Inactive - Archived') AND r.rating_type = 'Confirmed' AND total_qa_left > 0 AND r.cycle_start_date >= '$temp_year_start' AND r.cycle_start_date <= '$temp_year_end' ";

if($company!=''){
	$query.=" AND u.company_id = '$company' ";
}

if($state!=''){
	$query.=" AND u.state = '$state' ";
}

if($region!=''){
	$query.=" AND region_states.region = '$region' ";
}

$query.=" ORDER BY r.rater_id, r.cycle_end_date";

$result=mysql_query($query) or die(mysql_error());

$num=mysql_numrows($result);

//echo $query;

// Sampled tab

$query="SELECT  r.*, DATE_FORMAT(r.cycle_start_date, '%m/%d/%Y') AS start_date, u.firstname,u.lastname, u.state, region_states.region, ur.firstname rfi_firstname,ur.lastname rfi_lastname

FROM `rater_1qa_tracking` r 

LEFT OUTER JOIN users u ON r.rater_id = u.id  

LEFT OUTER JOIN users ur ON r.rfi_id = ur.id  

LEFT OUTER JOIN region_states ON u.`state` = region_states.state_short 

WHERE u.raterstatus NOT IN ('Terminated','Inactive - Archived') AND r.rating_type = 'Sampled' AND total_qa_left > 0 AND r.cycle_start_date >= '$temp_year_start' AND r.cycle_start_date <= '$temp_year_end' ";

if($company!=''){
	$query.=" AND u.company_id = '$company' ";
}

if($state!=''){
	$query.=" AND u.state = '$state' ";
}

if($region!=''){
	$query.=" AND region_states.region = '$region' ";
}

$query.=" ORDER BY r.rater_id, r.cycle_end_date";

$resultSampled=mysql_query($query) or die(mysql_error());

$num_Sampled=mysql_numrows($resultSampled);

// Completed tab

$query="SELECT r.*, DATE_FORMAT(r.cycle_start_date, '%m/%d/%Y') AS start_date, u.firstname,u.lastname, u.state, region_states.region, ur.firstname rfi_firstname,ur.lastname rfi_lastname  

FROM `rater_1qa_tracking` r 

LEFT OUTER JOIN users u ON r.rater_id = u.id  

LEFT OUTER JOIN users ur ON r.rfi_id = ur.id  

LEFT OUTER JOIN region_states ON u.`state` = region_states.state_short 

WHERE u.raterstatus NOT IN ('Terminated','Inactive - Archived') AND total_qa_left < 1 AND r.cycle_start_date >= '$temp_year_start' AND r.cycle_start_date <= '$temp_year_end' ";

if($company!=''){
	$query.=" AND u.company_id = '$company' ";
}

if($state!=''){
	$query.=" AND u.state = '$state' ";
}

if($region!=''){
	$query.=" AND region_states.region = '$region' ";
}

$query.=" ORDER BY r.rater_id, r.cycle_end_date";

$result_comp=mysql_query($query) or die(mysql_error());

$num_comp=mysql_numrows($result_comp);







// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("BER QAD")
        ->setLastModifiedBy("BER Admin")
        ->setTitle("3% QA Report")
        ->setSubject("3% QA Report")
        ->setDescription("3% QA Report")
        ->setKeywords("phpExcel")
        ->setCategory("3% QA");
        
// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Rater Name');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'RFI Name');
$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Year');
$objPHPExcel->getActiveSheet()->setCellValue('D1', 'Total Ratings');
$objPHPExcel->getActiveSheet()->setCellValue('E1', 'ES Ratings');
$objPHPExcel->getActiveSheet()->setCellValue('F1', 'QAs Required');
$objPHPExcel->getActiveSheet()->setCellValue('G1', 'ES QAs Required');
$objPHPExcel->getActiveSheet()->setCellValue('H1', 'QAs Done');
$objPHPExcel->getActiveSheet()->setCellValue('I1', 'ES QAs Done');
$objPHPExcel->getActiveSheet()->setCellValue('J1', 'QAs Left');
$objPHPExcel->getActiveSheet()->setCellValue('K1', 'ES QAs Left');
$objPHPExcel->getActiveSheet()->setCellValue('L1', 'ST');
$objPHPExcel->getActiveSheet()->setCellValue('M1', 'Reg');

$n=2;
//$qry= mysql_query("select * from tbl_agent ");
//while($d= mysql_fetch_array($qry)){
//for($i=1;$i<4;$i++){
while($row = mysql_fetch_array($result)) {
	if($row['rfi_id'] != 0) {

                    $onepercent_es_qa = 'N/A';

                    $total_es_qa_done = 'N/A';

                    $total_es_qa_left = 'N/A';

                } else {

                    $onepercent_es_qa = $row['onepercent_es_qa'];

                    $total_es_qa_done = $row['total_es_qa_done'];

                    $total_es_qa_left = $row['total_es_qa_left'];

                }

                $current_year= explode("/",$row['start_date']); 

                $current_year=$current_year[2];
				
				
 $objPHPExcel->getActiveSheet()->setCellValue('A'.$n, $row['firstname'].' '.$row['lastname']);
 $objPHPExcel->getActiveSheet()->setCellValue('B'.$n, $row['rfi_firstname'].' '.$row['rfi_lastname']);
 $objPHPExcel->getActiveSheet()->setCellValue('C'.$n, $current_year);
 $objPHPExcel->getActiveSheet()->setCellValue('D'.$n, $row['total_ratings']);
 $objPHPExcel->getActiveSheet()->setCellValue('E'.$n, $row['total_es_ratings']);
 $objPHPExcel->getActiveSheet()->setCellValue('F'.$n, $row['onepercent_qa']);
 $objPHPExcel->getActiveSheet()->setCellValue('G'.$n, $onepercent_es_qa);
 $objPHPExcel->getActiveSheet()->setCellValue('H'.$n, $row['total_qa_done']);
 $objPHPExcel->getActiveSheet()->setCellValue('I'.$n, $total_es_qa_done);
 $objPHPExcel->getActiveSheet()->setCellValue('J'.$n, $row['total_qa_left']);
 $objPHPExcel->getActiveSheet()->setCellValue('K'.$n, $total_es_qa_left);
 $objPHPExcel->getActiveSheet()->setCellValue('L'.$n, $row['state']);
 $objPHPExcel->getActiveSheet()->setCellValue('M'.$n, $row['region']);
   $n++;
}
//}

$header = 'A1:M1';
$objPHPExcel->getActiveSheet()->getStyle($header)->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('00cccccc');

$style = array(
    'font' => array('bold' => true,),
    'alignment' => array('horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,),
    );
$objPHPExcel->getActiveSheet()->getStyle($header)->applyFromArray($style);

/** @var PHPExcel_Cell $cell */
for ($col = ord('A'); $col <= ord('M'); $col++)
{
    $objPHPExcel->getActiveSheet()->getColumnDimension(chr($col))->setAutoSize(true);
}                
                
// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('Confirmed');




// Create a new worksheet, after the default sheet
$objPHPExcel->createSheet();

// Add some data to the second sheet, resembling some different data types
$objPHPExcel->setActiveSheetIndex(1);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Rater Name');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'RFI Name');
$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Year');
$objPHPExcel->getActiveSheet()->setCellValue('D1', 'Total Ratings');
$objPHPExcel->getActiveSheet()->setCellValue('E1', 'ES Ratings');
$objPHPExcel->getActiveSheet()->setCellValue('F1', 'QAs Required');
$objPHPExcel->getActiveSheet()->setCellValue('G1', 'ES QAs Required');
$objPHPExcel->getActiveSheet()->setCellValue('H1', 'QAs Done');
$objPHPExcel->getActiveSheet()->setCellValue('I1', 'ES QAs Done');
$objPHPExcel->getActiveSheet()->setCellValue('J1', 'QAs Left');
$objPHPExcel->getActiveSheet()->setCellValue('K1', 'ES QAs Left');
$objPHPExcel->getActiveSheet()->setCellValue('L1', 'ST');
$objPHPExcel->getActiveSheet()->setCellValue('M1', 'Reg');

$n=2;
//$qry=executeQuery("select * from tbl_technician ");
//while($d= mysql_fetch_array($qry)){
while($row = mysql_fetch_array($resultSampled)) { 

                if($row['rfi_id'] != 0) {

                    $onepercent_es_qa = 'N/A';

                    $total_es_qa_done = 'N/A';

                    $total_es_qa_left = 'N/A';

                } else {

                    $onepercent_es_qa = $row['onepercent_es_qa'];

                    $total_es_qa_done = $row['total_es_qa_done'];

                    $total_es_qa_left = $row['total_es_qa_left'];

                }
 $objPHPExcel->getActiveSheet()->setCellValue('A'.$n, $row['firstname'].' '.$row['lastname']);
 $objPHPExcel->getActiveSheet()->setCellValue('B'.$n, $row['rfi_firstname'].' '.$row['rfi_lastname']);
 $objPHPExcel->getActiveSheet()->setCellValue('C'.$n, $current_year);
 $objPHPExcel->getActiveSheet()->setCellValue('D'.$n, $row['total_ratings']);
 $objPHPExcel->getActiveSheet()->setCellValue('E'.$n, $row['total_es_ratings']);
 $objPHPExcel->getActiveSheet()->setCellValue('F'.$n, $row['onepercent_qa']);
 $objPHPExcel->getActiveSheet()->setCellValue('G'.$n, $onepercent_es_qa);
 $objPHPExcel->getActiveSheet()->setCellValue('H'.$n, $row['total_qa_done']);
 $objPHPExcel->getActiveSheet()->setCellValue('I'.$n, $total_es_qa_done);
 $objPHPExcel->getActiveSheet()->setCellValue('J'.$n, $row['total_qa_left']);
 $objPHPExcel->getActiveSheet()->setCellValue('K'.$n, $total_es_qa_left);
 $objPHPExcel->getActiveSheet()->setCellValue('L'.$n, $row['state']);
 $objPHPExcel->getActiveSheet()->setCellValue('M'.$n, $row['region']);
   $n++;
}
//}

$header = 'A1:M1';
$objPHPExcel->getActiveSheet()->getStyle($header)->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('00cccccc');

$style = array(
    'font' => array('bold' => true,),
    'alignment' => array('horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,),
    );
$objPHPExcel->getActiveSheet()->getStyle($header)->applyFromArray($style);

/** @var PHPExcel_Cell $cell */
for ($col = ord('A'); $col <= ord('M'); $col++)
{
    $objPHPExcel->getActiveSheet()->getColumnDimension(chr($col))->setAutoSize(true);
}       

// Rename 2nd sheet
$objPHPExcel->getActiveSheet()->setTitle('Sampled');






// Create a new worksheet, after the 2nd sheet
$objPHPExcel->createSheet();

// Add some data to the second sheet, resembling some different data types
$objPHPExcel->setActiveSheetIndex(2);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Rater Name');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'RFI Name');
$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Year');
$objPHPExcel->getActiveSheet()->setCellValue('D1', 'Total Ratings');
$objPHPExcel->getActiveSheet()->setCellValue('E1', 'ES Ratings');
$objPHPExcel->getActiveSheet()->setCellValue('F1', 'QAs Required');
$objPHPExcel->getActiveSheet()->setCellValue('G1', 'ES QAs Required');
$objPHPExcel->getActiveSheet()->setCellValue('H1', 'QAs Done');
$objPHPExcel->getActiveSheet()->setCellValue('I1', 'ES QAs Done');
$objPHPExcel->getActiveSheet()->setCellValue('J1', 'QAs Left');
$objPHPExcel->getActiveSheet()->setCellValue('K1', 'ES QAs Left');
$objPHPExcel->getActiveSheet()->setCellValue('L1', 'ST');
$objPHPExcel->getActiveSheet()->setCellValue('M1', 'Reg');

$n=2;
//$qry=executeQuery("select * from tbl_technician ");
//while($d= mysql_fetch_array($qry)){
while($row = mysql_fetch_array($result_comp)) { 

                    if($row['rfi_id'] != 0) {

                        $onepercent_es_qa = 'N/A';

                        $total_es_qa_done = 'N/A';

                        $total_es_qa_left = 'N/A';

                    } else {

                        $onepercent_es_qa = $row['onepercent_es_qa'];

                        $total_es_qa_done = $row['total_es_qa_done'];

                        $total_es_qa_left = $row['total_es_qa_left'];

                    }

 $objPHPExcel->getActiveSheet()->setCellValue('A'.$n, $row['firstname'].' '.$row['lastname']);
 $objPHPExcel->getActiveSheet()->setCellValue('B'.$n, $row['rfi_firstname'].' '.$row['rfi_lastname']);
 $objPHPExcel->getActiveSheet()->setCellValue('C'.$n, $current_year);
 $objPHPExcel->getActiveSheet()->setCellValue('D'.$n, $row['total_ratings']);
 $objPHPExcel->getActiveSheet()->setCellValue('E'.$n, $row['total_es_ratings']);
 $objPHPExcel->getActiveSheet()->setCellValue('F'.$n, $row['onepercent_qa']);
 $objPHPExcel->getActiveSheet()->setCellValue('G'.$n, $onepercent_es_qa);
 $objPHPExcel->getActiveSheet()->setCellValue('H'.$n, $row['total_qa_done']);
 $objPHPExcel->getActiveSheet()->setCellValue('I'.$n, $total_es_qa_done);
 $objPHPExcel->getActiveSheet()->setCellValue('J'.$n, $row['total_qa_left']);
 $objPHPExcel->getActiveSheet()->setCellValue('K'.$n, $total_es_qa_left);
 $objPHPExcel->getActiveSheet()->setCellValue('L'.$n, $row['state']);
 $objPHPExcel->getActiveSheet()->setCellValue('M'.$n, $row['region']);
   $n++;
}
//}

$header = 'A1:M1';
$objPHPExcel->getActiveSheet()->getStyle($header)->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('00cccccc');

$style = array(
    'font' => array('bold' => true,),
    'alignment' => array('horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,),
    );
$objPHPExcel->getActiveSheet()->getStyle($header)->applyFromArray($style);

/** @var PHPExcel_Cell $cell */
for ($col = ord('A'); $col <= ord('M'); $col++)
{
    $objPHPExcel->getActiveSheet()->getColumnDimension(chr($col))->setAutoSize(true);
}       

// Rename 3rd sheet
$objPHPExcel->getActiveSheet()->setTitle('Completed');



// Redirect output to a client’s web browser (Excel2007)
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="1% QA Reports.xlsx"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0

ob_end_clean();
$objWriter->save('php://output');
?>
