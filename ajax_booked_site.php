<?php
require_once("../lib/config.php");
require_once("../lib/constants.php");
require_once('../Classes/PHPExcel.php');
$logged_user_id = my_session('user_id');
if (isset($_REQUEST['source']) && ($_REQUEST['source'] == 'app')) {
    $logged_user_id = $_REQUEST['user_id'];
}
$action_type = $_REQUEST['action_type'];
$return_data  = array();
$name = $_REQUEST['name'];
$ref_id = '';

if ($action_type == "SEARCH_BY_DATE") {
    $site_name = $_REQUEST['site_name'];
    $today = date('Y-m-d');
    $query = "select sm.site_code,sm.site_name,bh.booking_no,bh.booking_from,bh.booking_to,bd.booking_amount from booking_detail bd 
    inner join site_master sm ON bd.site_id=sm.site_id
    inner join booking_header bh on bh.booking_id=bd.booking_id where
    sm.site_id ='" . $site_name . "' order by bh.booking_to desc";;
    // echo $query;die;
    $result = $db->query($query);
    while ($data = mysqli_fetch_assoc($result)) {
        $ret[] = $data;
    }
    $return_data  = array('status' => true, 'qry' => $query, 'booking_srch_list' => $ret);
    echo json_encode($return_data);
} elseif ($action_type == "SITE_LIST") {
    $query = "SELECT site_id,site_name,site_code from site_master";
    // echo $query;die;	
    $result = $db->query($query);

    while ($data = mysqli_fetch_assoc($result)) {
        $ret[] = $data;
    }
    $return_data  = array('status' => true, 'site_list' => $ret);
    echo json_encode($return_data);
} elseif ($action_type == "DOWNLOAD_EXCEL_SEARCH_BOOKED") {
    $site_name = $_REQUEST['site_name'];
    $today = date('Y-m-d');
    $query = "select sm.site_code,sm.site_name,bh.booking_no,bh.booking_from,bh.booking_to,bd.booking_amount from booking_detail bd 
    inner join site_master sm ON bd.site_id=sm.site_id
    inner join booking_header bh on bh.booking_id=bd.booking_id where
    sm.site_id ='" . $site_name . "' order by bh.booking_to desc";;
    // echo $query;die;
    $result = $db->query($query);
    $objPHPExcel = new PHPExcel();

    $phpColor = new PHPExcel_Style_Color();
    $phpColor->setRGB("000000");

    $style_cell = array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ) 
     ); 

    $objPHPExcel->getDefaultStyle()->getFont()
        ->setName('Book Antiqua')
        ->setSize(10)
        ->setBold(true)
        ->setColor($phpColor);
   
    $styleArray1 = array(
        'font'  => array(
            'bold'  => true,
            'color' => array('rgb' => '000000'),
            'size'  => 15,
            'name' => 'Times New Roman',
            'alignment' => 'center',
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        )

    );
    $objPHPExcel->setActiveSheetIndex(0);
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Booked Site List: site_name ' . $site_name);
    $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray($styleArray1);
    $objPHPExcel->getActiveSheet()->mergeCells('A1:F1');

    $objPHPExcel->getActiveSheet()->setCellValue('A2', 'SL No.');
    $objPHPExcel->getActiveSheet()->setCellValue('B2', 'Site Code');
    $objPHPExcel->getActiveSheet()->setCellValue('C2', 'Site Name');
    $objPHPExcel->getActiveSheet()->setCellValue('D2', 'Booking No');
    $objPHPExcel->getActiveSheet()->setCellValue('E2', 'Booking From');
    $objPHPExcel->getActiveSheet()->setCellValue('F2', 'Booking To');

    $objPHPExcel->getActiveSheet()->getStyle('A2:F2')->applyFromArray($style_cell);

    foreach (range('A', 'F') as $columnID) {
        $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
    }
    $rowCount = 3;
    $rowCount_new = 1;
    $existtempinid = array();
    while ($row = mysqli_fetch_assoc($result)) { //print_r($row);exit;  
        $objPHPExcel->getActiveSheet()->getRowDimension($rowCount_new)->setRowHeight(-1);


        $objPHPExcel->getActiveSheet()->setCellValue('A' . $rowCount, $rowCount_new++);
        $objPHPExcel->getActiveSheet()->setCellValue('B' . $rowCount, $row['site_code']);
        $objPHPExcel->getActiveSheet()->setCellValue('C' . $rowCount, $row['site_name']);
        $objPHPExcel->getActiveSheet()->setCellValue('D' . $rowCount, $row['booking_no']);
        $objPHPExcel->getActiveSheet()->setCellValue('E' . $rowCount, $row['booking_from']);
        $objPHPExcel->getActiveSheet()->setCellValue('F' . $rowCount, $row['booking_to']);
        $rowCount++;
    }
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    ob_start();
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="' . "disbursed_report" . date('jS-F-y H-i-s') . ".xlsx" . '"');
    header('Cache-Control: max-age=0');
    $objWriter->save("php://output");
    $xlsData = ob_get_contents();
    ob_end_clean();

    $file_name = 'List_Of_Booked_Sites' . $today;
    $return_data =  array(
        'status' => true, 'file_name' => $file_name,
        'file' => "data:application/vnd.ms-excel;base64," . base64_encode($xlsData)
    );
    echo json_encode($return_data);
    exit;
}
