<?php

date_default_timezone_set('America/New_York');

include 'vendor/autoload.php';

use Goutte\Client;
use lib\KarmaTek\Stocker;

//$companies = ["VZ"];
$companies = file('companies.txt', FILE_IGNORE_NEW_LINES);

$columns = range('A','E');
$row = '1';

$client = new Client();
$objPHPExcel = new PHPExcel();

// Set Excel properties
$objPHPExcel->getProperties()->setCreator("Stocker App");
$objPHPExcel->getProperties()->setLastModifiedBy("Stocker App");
$objPHPExcel->getProperties()->setTitle("An app to track stock data (Title)");
$objPHPExcel->getProperties()->setSubject("An app to track stock data (Subject)");
$objPHPExcel->getProperties()->setDescription("An app to track stock data (Description)");
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setTitle('Leveraged Free Cash Flow');
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(1);
$objPHPExcel->getActiveSheet()->setTitle('Cash & Equivalents');
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(2);
$objPHPExcel->getActiveSheet()->setTitle('Debt');
$objPHPExcel->setActiveSheetIndex(0);

// Start App
print "Getting Quarterly Headers....\n";

try {

    // Create Crawler Object
    $crawler = $client->request('GET', 'http://finance.yahoo.com/q/cf?s=VZ');

} catch (Exception $e) {

    print "Uh oh! Somthin' went wrong getting the web page:\n";
    var_dump($e);
    die;

}

// Filter for parent table
if ($parentTable = $crawler->filter('table.yfnc_tabledata1 td table'))
{

    // Create quarterly headers
    $qtlyHeaders = [];
    $parentTable->filter('tr')->eq(0)->children()->each(function ($node) use (&$qtlyHeaders) {
        $qtlyHeaders[] = $node->text();
    });

    $qtlyHeaders[] = 'Company';

    // Generate PHPExcel data
    for ($j=0;$j<=2;$j++)
    {

        $objPHPExcel->setActiveSheetIndex($j);

        for ($i=0;$i<count($columns);$i++)
        {
            $objPHPExcel->getActiveSheet()->SetCellValue($columns[$i] . $row, $qtlyHeaders[$i]);
        }

    }

    $row++;

} else {

    print "Uh oh! Somthin' went wrong filtering the web page:\n";
    die;

}

foreach ($companies as $company)
{
    print "Processing $company....\n";

    try {

        // Create Crawler Object
        $crawler = $client->request('GET', 'http://finance.yahoo.com/q/cf?s=' . $company);

    } catch (Exception $e) {

        print "Uh oh! Somthin' went wrong getting the web page:\n";
        var_dump($e);
        die;

    }

    // Filter for parent table
    if ($parentTable = $crawler->filter('table.yfnc_tabledata1 td table'))
    {

        $objPHPExcel->setActiveSheetIndex(0);

        Stocker::LevFreeCash($parentTable,$objPHPExcel,$company,$columns,$row);

    }

    try {

        // Create Crawler Object
        $crawler = $client->request('GET', 'http://finance.yahoo.com/q/bs?s=' . $company);

    } catch (Exception $e) {

        print "Uh oh! Somthin' went wrong getting the web page:\n";
        var_dump($e);
        die;

    }

    // Filter for parent table
    if ($parentTable = $crawler->filter('table.yfnc_tabledata1 td table'))
    {

        $objPHPExcel->setActiveSheetIndex(1);
        Stocker::CashAndEquiv($parentTable,$objPHPExcel,$company,$columns,$row);

        $objPHPExcel->setActiveSheetIndex(2);
        Stocker::Debt($parentTable,$objPHPExcel,$company,$columns,$row);

    }

    $row++;

}
// Write PhpExcel Object and save
$objPHPExcel->setActiveSheetIndex(0);
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save("stocker.xlsx");
