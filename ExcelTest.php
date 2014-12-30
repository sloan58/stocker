<?php

date_default_timezone_set('America/New_York');

include 'vendor/autoload.php';

//Set column range and starting row

$objPHPExcel = PHPExcel_IOFactory::load("SAS.xlsx");
$objPHPExcel->setActiveSheetIndex(0);
$highestRow = $objPHPExcel->getActiveSheet()->getHighestRow();
$highestCol = $objPHPExcel->getActiveSheet()->getHighestColumn();
$highestColIndex = PHPExcel_Cell::columnIndexFromString($highestCol);
for ($i=1; $i<=$highestRow;$i++)
{

    $val = $objPHPExcel->getActiveSheet()->getCell('A' . $i)->getValue();
    print "A" . $i . ": $val\n";

}

print $highestCol . "\n";
print $highestColIndex . "\n";
print PHPExcel_Cell::stringFromColumnIndex($highestColIndex);

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
