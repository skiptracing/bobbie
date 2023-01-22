<?php

ini_set('display_errors', 1);

require 'Datei.php';

$configArr = [
    ['SELLER_ID', 'config', '444'],
    ['EXPORT_API_KEY', 'config', 'youDeveloper'],
    ['SELLER_PREFIX', 'config', 'youDev'],
    ['KEY4', 'config', 'val4'],
    ['KEY5', 'config', 'val5']
];

$fileOutput = Datei::Modifizieren("Product-Upload.xlsm", $configArr);

if ($fileOutput) {
   
    Datei::Downloaden($fileOutput);
    
}
