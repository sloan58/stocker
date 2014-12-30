<?php

date_default_timezone_set('America/New_York');

include 'vendor/autoload.php';

//Set column range and starting row

$objSAS_PHPExcel = PHPExcel_IOFactory::load("SAS.xlsx");
$objSAS_PHPExcel->setActiveSheetIndex(0);
$highestRow = $objSAS_PHPExcel->getActiveSheet()->getHighestRow();
$highestCol = $objSAS_PHPExcel->getActiveSheet()->getHighestColumn();
$highestColIndex = PHPExcel_Cell::columnIndexFromString($highestCol);

for ($i=1; $i<=$highestColIndex;$i++)
{

    $letter = PHPExcel_Cell::stringFromColumnIndex($i);

    $val = $objSAS_PHPExcel->getActiveSheet()->getCell($letter . '1')->getValue();

    if ($val == "LFCF")
    {

        print "$val\n";

    }

}

//print $highestCol . "\n";
//print $highestColIndex . "\n";
//print PHPExcel_Cell::stringFromColumnIndex($highestColIndex);

$objWriter = PHPExcel_IOFactory::createWriter($objSAS_PHPExcel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
