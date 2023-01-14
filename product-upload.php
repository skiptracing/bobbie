<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;

$fileInput = "Product-Upload.xlsm";

$configArr = [
                 ['SELLER_ID', 'config', '444'], 
                 ['EXPORT_API_KEY', 'config', 'youDeveloper'], 
                 ['SELLER_PREFIX', 'config', 'youDev'], 
                 ['KEY4', 'config', 'val4'], 
                 ['KEY5', 'config', 'val5']
             ];

function DateiModifizieren($fileInput, $configArr)
{
    $reader = new XlsxReader();
    $spreadsheet = $reader->load($fileInput);
    $sheet = $spreadsheet->getSheetByName('Hilfe!');
    
    $seller_prefix = $configArr[2][2];
    $version = $sheet->getCell('I2')->getValue();
    $fileOutput = "Product-Upload-" . $seller_prefix . "-" . $version . ".xlsx";

    $sheet = $spreadsheet->getSheetByName('Werte');
    $sheet->fromArray($configArr, NULL, 'A2');

    $writer = new XlsxWriter($spreadsheet);
    $writer->setPreCalculateFormulas(false);
    $writer->save($fileOutput);

    return $fileOutput;
}

function DateiDownloaden($fileOutput)
{
    header('Content-Description: File Transfer');
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment; filename="' . basename($fileOutput) . '"');
    header('Expires: 0');
    header('Cache-Control: must-revalidate');
    header('Pragma: public');
    header('Content-Length: ' . filesize($fileOutput));
    
    readfile($fileOutput);
}

$fileOutput = DateiModifizieren($fileInput, $configArr);

DateiDownloaden($fileOutput);


