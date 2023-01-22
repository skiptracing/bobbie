<?php

require_once 'Xlsm.php';

final class Datei {

    public static function Modifizieren($fileInput, $configArr) {

        $outFile = '';

        try {

            $productXlsm = new Xlsm($fileInput);

            //add values to excel
            for ($i = 0; $i < 5; $i++) {
                for ($j = 0; $j < 3; $j++) {
                    $productXlsm->setCellValue("Werte", $i, $j, $configArr[$i][$j]);
                }
            }

            $ver = $productXlsm->getCellValue("Hilfe!", 1, 8);

            $refArr = explode(".", $fileInput);
            $outFile = $refArr[0] . "-" . $configArr[2][2] . "-" . $ver . ".xlsm";

            $productXlsm->saveXml();

            $productXlsm->saveFile($outFile);
            
        } catch (Exception $ex) {
            
        }

        return $outFile;
    }

    public static function Downloaden($fileOutput) {

        header('Content-Description: File Transfer');
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename="' . basename($fileOutput) . '"');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Content-Length: ' . filesize($fileOutput));

        readfile($fileOutput);
    }

}
