<?php

date_default_timezone_set('America/New_York');

include 'vendor/autoload.php';

use Goutte\Client;
use lib\KarmaTek\Stocker;

// Open a file with the list of companies
$companies = file('companies.txt', FILE_IGNORE_NEW_LINES);

//Set column range and starting row
$columns = range('A','F');
$row = '1';

// Create Goutte client
$client = new Client();

// Create PhpExcel Object
$objPHPExcel = new PHPExcel();
$objPHPExcel->getActiveSheet()->setCellValueExplicit();

// Set excel worksheet high level properties
$objPHPExcel->getProperties()->setCreator("Stocker App");
$objPHPExcel->getProperties()->setLastModifiedBy("Stocker App");
$objPHPExcel->getProperties()->setTitle("An app to track stock data (Title)");
$objPHPExcel->getProperties()->setSubject("An app to track stock data (Subject)");
$objPHPExcel->getProperties()->setDescription("An app to track stock data (Description)");

// Set title 'Leveraged Free Cash Flow' for tab 1 (tab 1 is created by default)
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setTitle('Leveraged Free Cash Flow');

// Create tab 2 and set title as 'Cash & Equivalents'
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(1);
$objPHPExcel->getActiveSheet()->setTitle('Cash & Equivalents');

// Create tab 3 and set title as 'Debt'
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(2);
$objPHPExcel->getActiveSheet()->setTitle('Debt');

// Create tab 4 and set title as 'Totals'.  Create headers for this tab as well
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(3);
$objPHPExcel->getActiveSheet()->setTitle('Totals');
$objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Company');
$objPHPExcel->getActiveSheet()->SetCellValue('B1', 'LFCF');
$objPHPExcel->getActiveSheet()->SetCellValue('C1', 'Cash and Equiv');
$objPHPExcel->getActiveSheet()->SetCellValue('D1', 'Debt');

// Set handle back to tab 1
$objPHPExcel->setActiveSheetIndex(0);

//Open SAS Spreadsheet and set active sheet to zero
$objReader = new PHPExcel_Reader_Excel2007();
$objSAS_PHPExcel = $objReader->load("SAS.xlsx");
$objSAS_PHPExcel->setActiveSheetIndex(0);

//Get highest row nad column information
$highestRow = $objSAS_PHPExcel->getActiveSheet()->getHighestRow();
$highestCol = $objSAS_PHPExcel->getActiveSheet()->getHighestColumn();
$highestColIndex = PHPExcel_Cell::columnIndexFromString($highestCol);

// Find where headers exist in Row 1
for ($i=1; $i<=$highestColIndex;$i++)
{
    $letter = PHPExcel_Cell::stringFromColumnIndex($i);

    $val = $objSAS_PHPExcel->getActiveSheet()->getCell($letter . '1')->getValue();

    switch($val)
    {

        case "LFCF":
            $lfcf_header = $letter;
            break;

        case "C&E":
            $ce_header = $letter;
            break;

        case "Debt":
            $debt_header = $letter;
            break;

    }

}

// Start App
print "Getting Quarterly Headers....\n";

try {

    // GET request to obtain quarterly headers.  Using VZ here since I thought it would be likely to always have that data.
    $crawler = $client->request('GET', 'http://finance.yahoo.com/q/cf?s=VZ');

} catch (Exception $e) {

    //  Error occurred GETing the URL
    print "Uh oh! Somthin' went wrong getting the web page:\n";
    var_dump($e);
    die;

}

// Filter for the table where interesting data begins
if ($parentTable = $crawler->filter('table.yfnc_tabledata1 td table'))
{

    // Create quarterly headers by crawling HTML table
    $qtlyHeaders = [];
    $parentTable->filter('tr')->eq(0)->children()->each(function ($node) use (&$qtlyHeaders) {
        $qtlyHeaders[] = $node->text();
    });

    // Set last column header to 'Company'
    $qtlyHeaders[] = 'Total';

    // Generate PHPExcel data for column headers on each tab
    for ($j=0;$j<=2;$j++)
    {

        $objPHPExcel->setActiveSheetIndex($j);

        for ($i=0;$i<count($columns);$i++)
        {
            $objPHPExcel->getActiveSheet()->SetCellValue($columns[$i] . $row, $qtlyHeaders[$i]);
        }

    }

    // Increment $row
    $row++;

} else {

    // There was a problem filtering for the headers
    print "Uh oh! Somthin' went wrong filtering the web page:\n";
    die;

}

// Begin iterating companies from the input file
for ($i=2; $i<=$highestRow;$i++)
//for ($i=2; $i<=3;$i++)
{
    //  cli status
    $company = $objSAS_PHPExcel->getActiveSheet()->getCell('A' . $i)->getValue();
    print "Processing $company....\n";

    try {

        // GET the company data from Yahoo!
        $crawler = $client->request('GET', 'http://finance.yahoo.com/q/cf?s=' . $company);

    } catch (Exception $e) {

        //  Error occurred GETing the URL
        print "Uh oh! Somthin' went wrong getting the web page:\n";
        var_dump($e);
        die;

    }

    // Filter for the table where interesting data begins
    if ($parentTable = $crawler->filter('table.yfnc_tabledata1 td table'))
    {

        // Set active sheet to 'Leveraged Free Cash Flow'
        $objPHPExcel->setActiveSheetIndex(0);

        // Run Stocker method to generate 'Leveraged Free Cash Flow'
        $lfcf_total = Stocker::LevFreeCash($parentTable,$objPHPExcel,$company,$columns,$row);

        // Update SAS Spreadsheet with total 'Leveraged Free Cash Flow'
        $col = PHPExcel_Cell::stringFromColumnIndex($highestColIndex);
        $objSAS_PHPExcel->getActiveSheet()->SetCellValue($lfcf_header . $i, $lfcf_total == 0 ? "No Data Available" : $lfcf_total);

    }

    try {

        // GET the company data from Yahoo!
        $crawler = $client->request('GET', 'http://finance.yahoo.com/q/bs?s=' . $company);

    } catch (Exception $e) {

        //  Error occurred GETing the URL
        print "Uh oh! Somthin' went wrong getting the web page:\n";
        var_dump($e);
        die;

    }

    // Filter for the table where interesting data begins
    if ($parentTable = $crawler->filter('table.yfnc_tabledata1 td table'))
    {

        // Set active sheet to 'Cash & Equivalents'
        $objPHPExcel->setActiveSheetIndex(1);

        // Run Stocker method to generate 'Cash & Equivalents'
        $ce_total = Stocker::CashAndEquiv($parentTable,$objPHPExcel,$company,$columns,$row);

        // Update SAS Spreadsheet with total 'Cash & Equivalents'
        $col++;
        $objSAS_PHPExcel->getActiveSheet()->SetCellValue($ce_header . $i, $ce_total);

        // Set active sheet to 'Debt'
        $objPHPExcel->setActiveSheetIndex(2);

        // Run Stocker method to generate 'Debt'
        $debt_total = Stocker::Debt($parentTable,$objPHPExcel,$company,$columns,$row);

        // Update SAS Spreadsheet with total 'Debt'
        $col++;
        $objSAS_PHPExcel->getActiveSheet()->SetCellValue($debt_header . $i, $debt_total);

    }

    // Increment $row
    $row++;

}
// Create PhpExcel Writer Object containing all sheet data, then save
$objPHPExcel->setActiveSheetIndex(0);
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save("stocker.xlsx");

$objWriter_SAS = new PHPExcel_Writer_Excel2007($objSAS_PHPExcel);
$objWriter_SAS->save("SAS-Stocker.xlsx");