<?php 
	
	try{
		
		$dbh = new pdo( 'mysql:host=localhost:3306;dbname=emp;charset=utf8',
				'root',
				'password',
				array(PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION));
	}catch(PDOException $ex){
		die(json_encode(array('outcome' => false, 'message' => 'Unable to connect')));
	}
	
	
	require '/vendor/autoload.php';
	
	error_reporting(E_ALL & ~E_NOTICE & ~E_WARNING & ~E_DEPRECATED);
	
	ini_set("display_errors", 1);

	use PhpOffice\PhpSpreadsheet\Spreadsheet;
	use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
	
	$spreadsheet = new Spreadsheet();
	
	
	$db_name = 'spmed';
	$create_data	=	date('Y-m-d');
	
	$spreadsheet->setActiveSheetIndex(0);
	$spreadsheet->getActiveSheet()->setTitle('Table List');
	
	$sheet = $spreadsheet->getActiveSheet()->mergeCells('B2:E2');
	$sheet->getStyle('B2:E2')->getAlignment()->setHorizontal('center');
	$sheet = $spreadsheet->getActiveSheet()->setCellValue('B2', 'TableLayout List');
	
	$sheet = $spreadsheet->getActiveSheet()->setCellValue('B3', 'Number');
	$sheet = $spreadsheet->getActiveSheet()->setCellValue('C3', 'Table ID');
	$sheet = $spreadsheet->getActiveSheet()->setCellValue('D3', 'Table Comment');
	$sheet = $spreadsheet->getActiveSheet()->setCellValue('E3', 'Etc');
	
	$sheet->getStyle('B')->getAlignment()->setHorizontal('center');
	$sheet->getStyle('C3')->getAlignment()->setHorizontal('center');
	$sheet->getStyle('D3')->getAlignment()->setHorizontal('center');
	$sheet->getStyle('E3')->getAlignment()->setHorizontal('center');
	
	$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(9);
	$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(21);
	$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(48);
	$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(15);
	
	$sql = 'SELECT table_name, table_comment FROM information_schema.tables
			WHERE table_schema = :DB';
	$rs = $dbh->prepare($sql);
	$rs->bindValue(':DB', $db_name);
	$rs->execute();
	$tb_list= $rs->fetchall();
	
	
	$num = 1;
	$index = 4;
	
	foreach ($tb_list as $tb_row){
		$sheet = $spreadsheet->getActiveSheet()->setCellValue('B'.$index, $num);
		$sheet = $spreadsheet->getActiveSheet()->setCellValue('C'.$index, $tb_row['table_name']);
		$sheet = $spreadsheet->getActiveSheet()->setCellValue('D'.$index, $tb_row['table_comment']);
		$sheet = $spreadsheet->getActiveSheet()->setCellValue('E'.$index, '');
				
		$num++;
		$index++;
	}
	
	$styleArray = [
			'borders' => [
					'allBorders' => [
							'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
							'color' => ['argb' => 'FF000000'],
					],
			],
	];
	
	
	$header_styleArray = [
			'fill' => [
					'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
					'startColor' => [
							'argb' => 'd1ebff',
					],
			],
			
	];
	
	$title_styleArray = [
			'font' => [
					'bold' => true,
			],
			'fill' => [
					'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
					'startColor' => [
							'argb' => 'd2ffd1',
					],
			],
	];
	
	$spreadsheet->getActiveSheet()->getStyle('B2:E'.($index-1))->applyFromArray($styleArray);
	
	$sql	= 'SELECT table_name, table_comment FROM information_schema.tables
				WHERE table_schema = :database_name';
	
	$rs = $dbh->prepare($sql);
	$rs->bindValue(':database_name', $db_name);
	$rs->execute();
	$tb_list= $rs->fetchall();
	
	$i = 1;
	
	
	foreach ($tb_list as $tb_row) {
		

		$myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, $tb_row['table_name']);
		
		$spreadsheet->addSheet($myWorkSheet, $i);
		
		$sheetIndex = $spreadsheet->getIndex($spreadsheet->getSheetByName($tb_row['table_name']));
		
 		$spreadsheet->setActiveSheetIndex($sheetIndex);
 		$sheet = $spreadsheet->getActiveSheet()->mergeCells('B2:B5');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('B2', 'Info');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('C2', 'Project');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('C3', 'Document');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('C4', 'Table_Name');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('C5', 'Writer');
 		
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('F4', 'Table_Comment');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('F5', 'Date');
 		
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('B6', 'No');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('C6', 'Column Name');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('D6', 'Type');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('E6', 'Size');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('F6', 'Column Key');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('G6', 'Comment');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('H6', 'Remarks');
 		
 		$sheet = $spreadsheet->getActiveSheet()->mergeCells('D4:E4');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('D4', $tb_row['table_name']);
 		
 		$sheet = $spreadsheet->getActiveSheet()->mergeCells('D2:H2');
 		$sheet = $spreadsheet->getActiveSheet()->mergeCells('D3:H3');
 		$sheet = $spreadsheet->getActiveSheet()->mergeCells('D5:E5');
 		
 		$sheet = $spreadsheet->getActiveSheet()->mergeCells('G4:H4');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('G4', $tb_row['table_comment']);
		
 		$sheet = $spreadsheet->getActiveSheet()->mergeCells('G5:H5');
 		$sheet = $spreadsheet->getActiveSheet()->setCellValue('G5', $create_data);
 		
 		$spreadsheet->getActiveSheet()->getStyle('B2:C5')->applyFromArray($header_styleArray);
 		$spreadsheet->getActiveSheet()->getStyle('F4:F5')->applyFromArray($header_styleArray);
 		$spreadsheet->getActiveSheet()->getStyle('B6:H6')->applyFromArray($title_styleArray);
 		
 		$spreadsheet->getActiveSheet()->getStyle('B2:H6')
 		->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
		
 		$sheet->getStyle('B2:H6')->getAlignment()->setHorizontal('center');
 		
 		
 		
 		$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(9);
 		$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(35);
 		$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(13);
 		$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(9);
 		$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);
 		$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(30);
 		$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);
 		
 		$sql = "SELECT UPPER(COLUMN_NAME) as COLUMN_NAME, UPPER(DATA_TYPE) as DATA_TYPE, 
									IFNULL(CHARACTER_MAXIMUM_LENGTH,NUMERIC_PRECISION+1) as LENGTH , 
									CONCAT(
									CASE COLUMN_KEY
									WHEN 'PRI' THEN 'PK'
									WHEN 'MUL' THEN 'INDEX'
									ELSE ''
									END ,IF(IS_NULLABLE='NO','NN','')) as COLUMN_TYPE, COLUMN_COMMENT
									FROM
									information_schema.columns
									WHERE
									table_schema = :database_name AND table_name = :tb_name";
 		
 		$rs = $dbh->prepare($sql);
 		$rs->bindValue(':database_name', $db_name);
 		$rs->bindValue(':tb_name', $tb_row['table_name']);
 		$rs->execute();
 		$column_list= $rs->fetchall();
 		
 		
 		$j = 7;
 		$num = 1;
 		
 		foreach ($column_list as $column_row){
 			
 			$sheet = $spreadsheet->getActiveSheet()->setCellValue('B'.$j, $num++);
 			$sheet = $spreadsheet->getActiveSheet()->setCellValue('C'.$j, $column_row['COLUMN_NAME']);
 			$sheet = $spreadsheet->getActiveSheet()->setCellValue('D'.$j, $column_row['DATA_TYPE']);
 			$sheet = $spreadsheet->getActiveSheet()->setCellValue('E'.$j, $column_row['LENGTH']);
 			$sheet = $spreadsheet->getActiveSheet()->setCellValue('F'.$j, $column_row['COLUMN_TYPE']);
 			$sheet = $spreadsheet->getActiveSheet()->setCellValue('G'.$j, $column_row['COLUMN_COMMENT']);
 			$sheet = $spreadsheet->getActiveSheet()->setCellValue('H'.$j, '');
 			
 			$j++;
 			
 		}
 		
 		$spreadsheet->getActiveSheet()->getStyle('B2:H'.($j-1))->applyFromArray($styleArray);
 		$sheet->getStyle('B7:B'.$j)->getAlignment()->setHorizontal('center');
 		$sheet->getStyle('D7:E'.$j)->getAlignment()->setHorizontal('center');
		
		$i++;
	}
	
		
	/*
	$sheetIndex = $spreadsheet->getIndex($spreadsheet->getSheetByName('Worksheet'));
	$spreadsheet->removeSheetByIndex($sheetIndex);
	*/

	
	// 칼럼 넓이 지정 - AUTO
	//foreach (range('B','H') as $col) { $spreadsheet->getActiveSheet()->getColumnDimension($col)->setAutoSize(true);}
	
	$date	=	date("Ymd");
	$xlsName = 'TableLayout_'.$date.'.xlsx';
	
	header('Content-Type: application/vnd.ms-excel');
	header('Content-Disposition: attachment;filename="'.$xlsName.'"');
	header('Cache-Control: max-age=0');
	
	$objWriter= new Xlsx($spreadsheet);
	$objWriter->save('php://output');

?>