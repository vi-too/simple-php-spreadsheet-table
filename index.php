<?php
require('vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet;
$listSheet = new Worksheet($spreadsheet, 'List');
$summarySheet = new Worksheet($spreadsheet, 'Summary');
$writer = new Xlsx($spreadsheet);
$spreadsheet->removeSheetByIndex(0);
$spreadsheet->addSheet($listSheet);
$spreadsheet->addSheet($summarySheet);

$listData = [
    [
        'ID' => 1,
        'Name' => 'grass',
        'Weekness' => 'fire'
    ],
    [
        'ID' => 2,
        'Name' => 'grass',
        'Weekness' => 'fire'
    ],
];

$summaryData = [
    [
        [
            'Month' => 'Jan',
            'Count' => 3,
        ],
        [
            'Month' => 'Feb',
            'Count' => 4,
        ],
    ],
    [
        [
            'Menthed' => 'Jan',
            'Count' => 3,
        ],
        [
            'Month' => 'Feb',
            'Count' => 4,
        ],
        [
            'Month' => 'Feb',
            'Count' => 4,
        ],
        [
            'Month' => 'Feb',
            'Count' => 4,
        ],
    ],
    [
        [
            'Marlked' => 'Jan',
            'Count' => 3,
        ],
        [
            'Marlked' => 'Feb',
            'Count' => 4,
        ],
    ],

];

function writeListSheet(Worksheet $sheet, array $data)
{
    $table = new Table($sheet, $data);
    $table->write();
}

function writeSummarySheet(Worksheet $sheet, array $groups)
{
    $perRow = 2;
    $gutterColumn = 1;
    $gutterRow = 2;
    $column = 1;
    $row = 1;
    $maxRow = 0;

    foreach ($groups as $groupIndex => $data) {
        $table = new Table($sheet, $data, $sheet->getCellByColumnAndRow($column, $row)->getCoordinate());

        $table->write();
        // Adjust max row
        if ($table->getLastRow() > $maxRow) {
            $maxRow = $table->getLastRow();
        }

        // Move to next eligible row and reset column
        if (($groupIndex + 1) % $perRow === 0 && $groupIndex !== 0) {
            $row = ($maxRow + $gutterRow) + 1;
            $column = 1;
        } else {
            $column = ($table->getLastColumnIndex() + $gutterColumn) + 1;
        }
    }
}

writeListSheet($listSheet, $listData);
writeSummarySheet($summarySheet, $summaryData);

function dd(...$args)
{
    echo '<pre>' . json_encode($args, JSON_PRETTY_PRINT) . '</pre>';
    die();
}

function logger($message, $data)
{
    if (is_array($data)) {
        $message .= ' ' . json_encode($data);
    }

    file_put_contents('./main.log', $message . PHP_EOL, FILE_APPEND);
}


header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename=myfilename.xlsx');
$writer->save('php://output');
die();
