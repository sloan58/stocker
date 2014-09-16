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
            $tcffoa[] = trim(str_replace(['(',')',',','-'],'',$node->text()));
        });


        // Create Capital Expenditures
        $capex = [];

        if ($parentTable->filter('tr')->eq(14)->filter('td')->count() == 5)
        {

            $parentTable->filter('tr')->eq(14)->filter('td')->each(function ($node) use (&$capex,&$company) {

                $capex[] = trim(str_replace(['(',')',',','-'],'',$node->text()));

            });

        }

        // Create Leveraged Free Cash Flow
        $lfcf = [];
        for ($i=0;$i<count($columns);$i++)
        {
            if ($i == 0)
            {

                $lfcf[] = $company;

            } elseif (!isset($tcffoa[$i]) || !isset($capex[$i])) {

                $lfcf[$i] = 'No Data Available';

            } else {

                $lfcf[$i] = $tcffoa[$i] - $capex[$i];

            }
        }

        // Write cell content for Leveraged Free Cash Flow
        for ($i=0;$i<count($columns);$i++)
        {
            $objPHPExcel->getActiveSheet()->SetCellValue($columns[$i] . $row, $lfcf[$i]);
        }

        return true;
    }

    public static function CashAndEquiv($parentTable,$objPHPExcel,$company,$columns,$row)
    {

        // Create Cash And Cash Equivalents
        $cce = [];

        $parentTable->filter('tr')->eq(4)->filter('td')->each(function ($node) use (&$cce) {
            $cce[] = trim(str_replace(['(',')',',','-'],'',$node->text()));
        });

        // Yahoo! buffer <td>
        array_shift($cce);

        // Create Short Term Investments
        $sti = [];

        $parentTable->filter('tr')->eq(5)->filter('td')->each(function ($node) use (&$sti) {
            $sti[] = trim(str_replace(['(',')',',','-'],'',$node->text()));
        });

        // Yahoo! buffer <td>
        array_shift($sti);

        // Create Cash & Equivalents
        $ce = [];
        for ($i=0;$i<count($columns);$i++)
        {

            if ($i == 0)
            {

                $ce[] = $company;

            } elseif (!isset($cce[$i]) || !isset($sti[$i])) {

                $ce[$i] = 'No Data Available';

            } else {

                $ce[$i] = $cce[$i] + $sti[$i];

            }

        }

        // Write cell content for Leveraged Free Cash Flow
        for ($i=0;$i<count($columns);$i++)
        {
            $objPHPExcel->getActiveSheet()->SetCellValue($columns[$i] . $row, $ce[$i]);
        }

        return true;

    }

    public static function Debt($parentTable,$objPHPExcel,$company,$columns,$row)
    {

        // Create Short/Current Long Term Debt
        $scltd = [];


        if ($parentTable->filter('tr')->eq(24)->filter('td')->count() == 6)
        {

            $parentTable->filter('tr')->eq(24)->filter('td')->each(function ($node) use (&$scltd) {
                $scltd[] = trim(str_replace(['(',')',',','-'],'',$node->text()));
            });

        }

        // Yahoo! buffer <td>
        array_shift($scltd);


        // Create Long Term Debt
        $ltd = [];

        if ($parentTable->filter('tr')->eq(28)->filter('td')->count() == 5)
        {

            $parentTable->filter('tr')->eq(28)->filter('td')->each(function ($node) use (&$ltd) {
                $ltd[] = trim(str_replace(['(',')',',','-'],'',$node->text()));
            });

        }


        // Create Debt
        $debt = [];
        for ($i=0;$i<count($columns);$i++)
        {

            if ($i == 0)
            {

                $debt[] = $company;

            } elseif (!isset($scltd[$i]) || !isset($ltd[$i])) {

                $debt[$i] = 'No Data Available';

            } else {

                $debt[$i] = $scltd[$i] + $ltd[$i];

            }

        }

        // Write cell content for Leveraged Free Cash Flow
        for ($i=0;$i<count($columns);$i++)
        {
            $objPHPExcel->getActiveSheet()->SetCellValue($columns[$i] . $row, $debt[$i]);
        }

        return true;

    }
} 