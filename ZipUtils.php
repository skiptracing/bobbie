<?php

final class ZipUtils {

    static public function folderToZip($folder, &$zipFile, $exclusiveLength) {

        $handle = opendir($folder);

        while (false !== $f = readdir($handle)) {

            if ($f != '.' && $f != '..') {

                $filePath = "$folder/$f";

                $localPath = substr($filePath, $exclusiveLength);

                if (is_file($filePath)) {

                    $zipFile->addFile($filePath, $localPath);
                } elseif (is_dir($filePath)) {

                    $zipFile->addEmptyDir($localPath);

                    self::folderToZip($filePath, $zipFile, $exclusiveLength);
                }
            }
        }

        closedir($handle);
    }

    static public function zipDir($sourcePath, $outZipPath) {

        $pathInfo = pathInfo($sourcePath);

        $z = new ZipArchive();

        $z->open($outZipPath, ZIPARCHIVE::CREATE);

        self::folderToZip($sourcePath, $z, strlen("$sourcePath") + 1);

        $z->close();
    }

    static public function unzipDir($sourcePath, $outDir) {
        $zip = new ZipArchive();
        if ($res = $zip->open($sourcePath)) {
            $zip->extractTo($outDir);
            $zip->close();
        }
    }

    static public function recurseRmdir($dir) {
        $files = array_diff(scandir($dir), array('.', '..'));
        foreach ($files as $file) {
            (is_dir("$dir/$file") && !is_link("$dir/$file")) ? self::recurseRmdir("$dir/$file") : unlink("$dir/$file");
        }
        return rmdir($dir);
    }

}
