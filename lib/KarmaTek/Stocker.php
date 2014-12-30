<?php
/**
 * Created by PhpStorm.
 * User: sloan58
 * Date: 9/3/14
 * Time: 7:41 PM
 */

namespace lib\KarmaTek;


class Stocker {

    public static function LevFreeCash($parentTable,$objPHPExcel,$company,$columns,$row)
    {

        // Create Total Cash Flow From Operating Activities
        $tcffoa = [];

        $parentTable->filter('tr')->eq(11)->filter('td strong')->each(function ($node) use (&$tcffoa) {
            $tcffoa[] = self::cleanStr($node->text());
        });


        // Create Capital Expenditures
        $capex = [];

        if ($parentTable->filter('tr')->eq(14)->filter('td')->count() == 5)
        {

            $parentTable->filter('tr')->eq(14)->filter('td')->each(function ($node) use (&$capex,&$company) {

                $capex[] = self::cleanStr($node->text());

            });

        }

        // Create Leveraged Free Cash Flow
        $lfcf = self::calculateQuarter($columns,$company,$tcffoa,$capex,'-');

        // Sum all 4 quarters
        $lfcf_total = self::sumArray($lfcf);

        // Write cell content for Leveraged Free Cash Flow
        self::writeData($objPHPExcel,$columns,$row,$lfcf,$lfcf_total);

        // Write totals in the 'Totals' tab
        self::writeTotalsData($objPHPExcel,'B',$row,$company,$lfcf_total);

        return $lfcf_total;

    }

    public static function CashAndEquiv($parentTable,$objPHPExcel,$company,$columns,$row)
    {

        // Create Cash And Cash Equivalents
        $cce = [];

        $parentTable->filter('tr')->eq(4)->filter('td')->each(function ($node) use (&$cce) {
            $cce[] = self::cleanStr($node->text());
        });

        // Yahoo! buffer <td>
        array_shift($cce);

        // Create Short Term Investments
        $sti = [];

        $parentTable->filter('tr')->eq(5)->filter('td')->each(function ($node) use (&$sti) {
            $sti[] = self::cleanStr($node->text());
        });

        // Yahoo! buffer <td>
        array_shift($sti);

        // Create Cash & Equivalents
        $ce = self::calculateQuarter($columns,$company,$cce,$sti);

        // Sum all 4 quarters
        $ce_total = self::sumArray($ce);

        // Write cell content for Cash and Equivalents
        self::writeData($objPHPExcel,$columns,$row,$ce,$ce_total);

        // Write totals in the 'Totals' tab
        self::writeTotalsData($objPHPExcel,'C',$row,$company,$ce_total);

        return $ce[1];

    }

    public static function Debt($parentTable,$objPHPExcel,$company,$columns,$row)
    {

        // Create Short/Current Long Term Debt
        $scltd = [];


        if ($parentTable->filter('tr')->eq(24)->filter('td')->count() == 6)
        {

            $parentTable->filter('tr')->eq(24)->filter('td')->each(function ($node) use (&$scltd) {
                $scltd[] = self::cleanStr($node->text());
            });

        }

        // Yahoo! buffer <td>
        array_shift($scltd);


        // Create Long Term Debt
        $ltd = [];

        if ($parentTable->filter('tr')->eq(28)->filter('td')->count() == 5)
        {

            $parentTable->filter('tr')->eq(28)->filter('td')->each(function ($node) use (&$ltd) {
                $ltd[] = self::cleanStr($node->text());
            });

        }


        // Create Debt
        $debt = self::calculateQuarter($columns,$company,$scltd,$ltd);

        // Sum all 4 quarters
        $debt_total = self::sumArray($debt);

        // Write cell content for Debt
        self::writeData($objPHPExcel,$columns,$row,$debt,$debt_total);

        // Write totals in the 'Totals' tab
        self::writeTotalsData($objPHPExcel,'D',$row,$company,$debt_total);

        return $debt[1];

    }

    /**
     * Private Functions
     *
     */
    private static function cleanStr($string)
    {
        return trim(str_replace(['(',')',',','-'],'',$string));
    }

    private static function calculateQuarter($columns,$company,$array1,$array2,$plus_minus = '+')
    {
        $results = [];
        for ($i=0;$i<count($columns);$i++)
        {

            if ($i == 0)
            {

                $results[] = $company;

            } elseif (!isset($array1[$i]) || !isset($array2[$i])) {

                $results[$i] = 'No Data Available';

            } else {

                switch($plus_minus)
                {

                    case '+':

                        $results[$i] = ($array1[$i] + $array2[$i]) * 1000;
                        break;

                    case '-':

                        $results[$i] = ($array1[$i] - $array2[$i]) * 1000;
                        break;
                }

            }
        }
        return $results;
    }

    private static function sumArray($array)
    {
        if (!is_numeric($array[1]) && !is_numeric($array[2]) && !is_numeric($array[3]) && !is_numeric($array[4]))
        {
            return "No Data Available";
        }

        return array_sum([$array[1],$array[2],$array[3],$array[4]]);
    }

    private static function writeData($objPHPExcel,$columns,$row,$array,$total)
    {
        // Write cell content for each quarter
        for ($i=0;$i<count($columns);$i++)
        {
            $objPHPExcel->getActiveSheet()->SetCellValue($columns[$i] . $row, $array[$i]);
        }

        // Write total for 4 quarters
        $objPHPExcel->getActiveSheet()->SetCellValue($columns[count($columns) - 1] . $row, $total);

        return true;
    }

    private static function writeTotalsData($objPHPExcel,$column,$row,$company,$total)
    {
        // Grab 'Totals' tab
        $objPHPExcel->setActiveSheetIndex(3);

        // Record the company name
        $objPHPExcel->getActiveSheet()->SetCellValue('A' . $row, $company);

        // Write total for 4 quarters
        $objPHPExcel->getActiveSheet()->SetCellValue($column . $row, $total);
    }

}