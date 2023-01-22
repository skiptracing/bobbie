<?php

require_once 'ZipUtils.php';

class Xlsm {

    public $unpackedDir;
    public $xlsxFile;
    public $zipFile;
    public $workbookPath;
    public $workbookDom;
    public $sharedStringsPath;
    public $sharedStringsDom;
    public $sheetArr;
    public $fileExt;

    function __construct($fileInput) {

        $refArr = explode(".", $fileInput);

        $this->xlsxFile = $fileInput;
        $this->unpackedDir = $refArr[0];
        $this->zipFile = $refArr[0] . ".zip";
        $this->fileExt = $refArr[1];

        $this->sharedStringsPath = $refArr[0] . "/xl/sharedStrings.xml";
        $this->workbookPath = $refArr[0] . "/xl/workbook.xml";

        copy($this->xlsxFile, $this->zipFile);
        ZipUtils::unzipDir($this->zipFile, $refArr[0]);

        //load sheetArr
        $handle = opendir($this->unpackedDir . "/xl/worksheets");
        while (($entry = readdir($handle)) !== FALSE) {
            if (strpos($entry, "heet") > 0) {
                $this->sheetArr[] = [$entry, NULL, NULL]; //[filename, domXml, changed]
            }
        }

        $this->workbookDom = $this->loadDomXml($this->workbookPath);
        $this->sharedStringsDom = $this->loadDomXml($this->sharedStringsPath);

        //var_dump($this->sheetArr);
    }

    function loadDomXml($filePath) {

        $fileContent = file_get_contents($filePath);

        $domDocument = new DOMDocument();
        $domDocument->loadXML($fileContent);

        return $domDocument;
    }

    function saveXml() {

        $this->sharedStringsDom->saveXML();
        $this->sharedStringsDom->save($this->sharedStringsPath);

        foreach ($this->sheetArr as $sheet) {
            if ($sheet[2] == 1) {
                $sheet[1]->saveXML();
                $sheet[1]->save($this->unpackedDir . "/xl/worksheets/" . $sheet[0]);
            }
        }
    }

    function loadSheetPos($sheetName) {

        $workbookDom = $this->loadDomXml($this->workbookPath);

        $workbookSheets = $workbookDom->getElementsByTagName("sheets");
        $sheetFile = NULL;
        foreach ($workbookSheets[0]->childNodes as $node) {
            if ($sheetName == $node->getAttribute("name")) {
                $rid = $node->getAttribute("r:id");
                $sheetFile = str_replace("rId", "sheet", $rid) . '.xml';
                break;
            }
        }

        $k = 0;
        foreach ($this->sheetArr as $sheet) {
            if ($sheet[0] == $sheetFile) {
                if ($sheet[1] == NULL) {
                    $this->sheetArr[$k][1] = $this->loadDomXml(
                            $this->unpackedDir . "/xl/worksheets/" . $sheetFile);
                }
                break;
            }
            $k++;
        }

        return $k;
    }

    function getCellValue($sheetName, $rowIndex, $colIndex) {

        $sheetPos = $this->loadSheetPos($sheetName);
        $sheetDom = $this->sheetArr[$sheetPos][1];

        $rowArr = $sheetDom->getElementsByTagName("row");

        $rowElement = $rowArr[$rowIndex];

        $key = intval($rowElement->childNodes[$colIndex]->nodeValue);

        $sharedStrings = $this->loadDomXml($this->sharedStringsPath);

        $siArr = $sharedStrings->getElementsByTagName("si");

        $value = $siArr[$key]->nodeValue;

        return $value;
    }

    function numToLettersExcel($n) {
        $n -= 1;
        for ($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n % 26 + 0x41) . $r;
        }
        return $r;
    }

    function leterrsToNumExcel($a) {

        $l = strlen($a);
        $n = 0;
        for ($i = 0; $i < $l; $i++) {
            $n = $n * 26 + ord($a[$i]) - 0x40;
        }

        return $n;
    }

    function setCellValue($sheetName, $rowIndex, $colIndex, $value) {
        //
        $t_rows = $this->sharedStringsDom->getElementsByTagName("t");
        $pos_t = count($t_rows) - 1;

        $sstNode = $this->sharedStringsDom->documentElement;
        $siNode = $this->sharedStringsDom->createElement("si");
        $tNode = $this->sharedStringsDom->createElement("t", $value);

        $siNode->appendChild($tNode);
        $sstNode->appendChild($siNode);

        //
        $colIdStr = $this->numToLettersExcel($colIndex + 1) . $rowIndex;

        $sheetPos = $this->loadSheetPos($sheetName);
        $sheetDom = $this->sheetArr[$sheetPos][1];

        $rowArr = $sheetDom->getElementsByTagName("row");
        if (count($rowArr)) {
            $k = 0;
            $found = false;
            foreach ($rowArr as $row) {
                if (intval($row->getAttribute("r")) === $rowIndex) {
                    $found = true;
                    break;
                }
                $k++;
            }

            if ($found) {
                $colArr = $rowArr[$k]->getElementsByTagName("c");
                if (count($colArr)) {
                    $k = 0;
                    $found = false;
                    foreach ($colArr as $col) {

                        if ($col->getAttribute("r") === $colIdStr) {
                            $found = true;
                            break;
                        }
                        $k++;
                    }
                    if ($found) {
                        $colArr[$k]->nodeValue = $value;
                    } else {
                        //create new col and put value
                    }
                }
            } else {
                //create new row
                //create new col and put value
            }
        }

        $this->sheetArr[$sheetPos][2] = 1;
    }

    function saveFile($outFile) {

        ZipUtils::zipDir($this->unpackedDir, $outFile);
        ZipUtils::recurseRmdir($this->unpackedDir);
        unlink($this->zipFile);
    }

}
