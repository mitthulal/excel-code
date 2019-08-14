<?php
require 'Classes/PHPExcel.php';
    error_reporting(E_ALL);
    ini_set('display_errors', 1);
   $servername = "localhost";
$username = "blallab_liveu";
$password = '3;.5a;v3&2{l';
$dbname = 'blallab_live';
    // Create connection
    $conn = new mysqli($servername, $username, $password, $dbname);
    // Check connection
    if ($conn->connect_error) {
        die("Connection failed: " . $conn->connect_error);
    } 
	
	//SELECT `type`, count(*) FROM `items` WHERE now()>`endScheduleTime` group by type order by count(*) desc
	
	$query = mysqli_query($conn,"SELECT tblbooking_packages.full_name,tblbooking_packages.mobile,tblbooking_packages.email,tblbooking_packages.address,tblbooking_packages.preferred_date,tblbooking_packages.preferred_time,tblproducts.product_name,tbllocations.location_name, tblbooking_packages.payment_status,tblbooking_packages.net_amount,tblbooking_packages.ip_address  FROM tblbooking_packages 
LEFT JOIN tblproducts ON tblproducts.id = tblbooking_packages.package_id 
LEFT JOIN tbllocations ON tbllocations.id = tblbooking_packages.location_id 
ORDER BY tblbooking_packages.created DESC"
	);
	
	$objPHPExcel = new PHPExcel();
   $objPHPExcel->getActiveSheet()->setTitle('Blal Lab Bookings');

   // Loop through the result set
    $rowNumber = 1;
    
    $objPHPExcel->getActiveSheet()->setCellValue('A1',"Full Name");
    $objPHPExcel->getActiveSheet()->setCellValue('B1',"Mobile");
    $objPHPExcel->getActiveSheet()->setCellValue('c1',"Email");
    $objPHPExcel->getActiveSheet()->setCellValue('D1',"Address");
    $objPHPExcel->getActiveSheet()->setCellValue('E1',"Preferred Date");
    $objPHPExcel->getActiveSheet()->setCellValue('F1',"Preferred Time");
    $objPHPExcel->getActiveSheet()->setCellValue('G1',"Product Name");
    $objPHPExcel->getActiveSheet()->setCellValue('H1',"Location");
    $objPHPExcel->getActiveSheet()->setCellValue('I1',"Payment Status");
    $objPHPExcel->getActiveSheet()->setCellValue('J1',"Net Amount");
    $objPHPExcel->getActiveSheet()->setCellValue('K1',"Ip Address");
    
    
    
    $rowNumber = 2;

//Heading of excel sheet->Product Name, Available Count, Schedule Start Hour, Schedule End Hour, Total Count Generate Today

	while($row = mysqli_fetch_array($query,MYSQLI_ASSOC))
	{  
       $col = 'A';
       foreach($row as $cell) {
          $objPHPExcel->getActiveSheet()->setCellValue($col.$rowNumber,$cell);
          
          $col++;
       }
       $rowNumber++;
   // Save as an Excel BIFF (xls) file
   
   $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
	}
   header('Content-Type: application/vnd.ms-excel');
   header('Content-Disposition: attachment;filename="myFile.xls"');
   header('Cache-Control: max-age=0');

   $objWriter->save('php://output');
   exit();
?>