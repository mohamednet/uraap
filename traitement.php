<?php 
$date = date("Y_m_d_H_i_s");
require_once './phpexcel/Classes/PHPExcel/IOFactory.php';
$target_dir = "files/";
$x = $_FILES["fileToUpload"]["name"];
$target_file = $target_dir . "upload/$x";
$uploadOk = 1;
if (move_uploaded_file($_FILES["fileToUpload"]["tmp_name"], $target_file)) {
    echo "The file ". basename( $_FILES["fileToUpload"]["name"]). " has been uploaded.";
} else {
    echo "Sorry, there was an error uploading your file.";
}

// Chargement du fichier Excel 
$objPHPExcel = PHPExcel_IOFactory::load("files/upload/$x");
 
/**
* récupération de la première feuille du fichier Excel
* @var PHPExcel_Worksheet $sheet
*/
$sheet = $objPHPExcel->getSheet(0);
 $content_array = array();
 
// On boucle sur les lignes
foreach($sheet->getRowIterator() as $row) {
 $temp_array = array();
   // On boucle sur les cellule de la ligne
   foreach ($row->getCellIterator() as $cell) {
      $temp_array[] = $cell->getValue();
   }
 $content_array[] = $temp_array;
}



$objPHPExcel = new PHPExcel(); // Create new PHPExcel object

$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', 'Order number')
            ->setCellValue('B1', 'sku')
            ->setCellValue('C1', 'Item Name')
            ->setCellValue('D1', 'Item Quantity')
            ->setCellValue('E1', 'Price')
            ->setCellValue('F1', ' ')
            ->setCellValue('J1', 'Shipping Name')
            ->setCellValue('H1', 'Shipping Address 1')
            ->setCellValue('I1', 'Shipping Address 2')
            ->setCellValue('G1', 'Shipping City')
            ->setCellValue('K1', 'Shipping Province')
            ->setCellValue('L1', 'Shipping Country')
            ->setCellValue('M1', 'Shipping Zip')
            ->setCellValue('N1', 'Shipping phone')
            ->setCellValue('O1', 'Alternative phone')
            ->setCellValue('P1', ' ')
            ->setCellValue('Q1', 'Discount')
            ->setCellValue('R1', 'Price*Qty')
            ->setCellValue('S1', 'Tracking number');
for($i=0;$i<count($content_array);$i++)
{
    if($i>=1 && $content_array[$i][0])
    {
    $objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('A'.($i+1), $content_array[$i][0])
            ->setCellValue('B'.($i+1), " ")
            ->setCellValue('C'.($i+1), $content_array[$i][2])
            ->setCellValue('D'.($i+1), $content_array[$i][1])
            ->setCellValue('E'.($i+1), " ")
            ->setCellValue('F'.($i+1), ' ')
            ->setCellValue('J'.($i+1), $content_array[$i][3])
            ->setCellValue('H'.($i+1), $content_array[$i][4])
            ->setCellValue('I'.($i+1), $content_array[$i][5])
            ->setCellValue('G'.($i+1), $content_array[$i][7])
            ->setCellValue('K'.($i+1), $content_array[$i][9])
            ->setCellValue('L'.($i+1), $content_array[$i][11])
            ->setCellValue('M'.($i+1), $content_array[$i][10])
            ->setCellValue('N'.($i+1), $content_array[$i][12])
            ->setCellValue('O'.($i+1), $content_array[$i][13])
            ->setCellValue('P'.($i+1), ' ')
            ->setCellValue('Q'.($i+1), 1)
            ->setCellValue('R'.($i+1), " ")
            ->setCellValue('S'.($i+1), " ");
}
}

$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save("files/download/$x");
echo "<a href ='files/download/$x'>download here </a>";
?>