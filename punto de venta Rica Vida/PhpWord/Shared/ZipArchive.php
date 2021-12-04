<?php


namespace PhpOffice\PhpWord\Shared;

use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\Settings;


class ZipArchive
{
    
    const CREATE    = 1; 
    const OVERWRITE = 8; 

   
    public $numFiles = 0;

   
    public $filename;

    
    private $tempDir;

   
    private $zip;

    
    private $usePclzip = true;

    public function __construct()
    {
        $this->usePclzip = (Settings::getZipClass() != 'ZipArchive');
        if ($this->usePclzip) {
            if (!defined('PCLZIP_TEMPORARY_DIR')) {
                define('PCLZIP_TEMPORARY_DIR', Settings::getTempDir() . '/');
            }
            require_once 'PCLZip/pclzip.lib.php';
        }
    }

   
    public function __call($function, $args)
    {
        $zipFunction = $function;
        if (!$this->usePclzip) {
            $zipObject = $this->zip;
        } else {
            $zipObject = $this;
            $zipFunction = "pclzip{$zipFunction}";
        }

        $result = false;
        if (method_exists($zipObject, $zipFunction)) {
            $result = @call_user_func_array(array($zipObject, $zipFunction), $args);
        }

        return $result;
    }

  
    public function open($filename, $flags = null)
    {
        $result = true;
        $this->filename = $filename;

        if (!$this->usePclzip) {
            $zip = new \ZipArchive();
            $result = $zip->open($this->filename, $flags);

            
            $this->numFiles = $zip->numFiles;
        } else {
            $zip = new \PclZip($this->filename);
            $this->tempDir = Settings::getTempDir();
            $this->numFiles = count($zip->listContent());
        }
        $this->zip = $zip;

        return $result;
    }

  
    public function close()
    {
        if (!$this->usePclzip) {
            if ($this->zip->close() === false) {
                throw new Exception("Could not close zip file {$this->filename}.");
            }
        }

        return true;
    }

    
    public function extractTo($destination, $entries = null)
    {
        if (!is_dir($destination)) {
            return false;
        }

        if (!$this->usePclzip) {
            return $this->zip->extractTo($destination, $entries);
        } else {
            return $this->pclzipExtractTo($destination, $entries);
        }
    }

    
    public function getFromName($filename)
    {
        if (!$this->usePclzip) {
            $contents = $this->zip->getFromName($filename);
            if ($contents === false) {
                $filename = substr($filename, 1);
                $contents = $this->zip->getFromName($filename);
            }
        } else {
            $contents = $this->pclzipGetFromName($filename);
        }

        return $contents;
    }

    
    public function pclzipAddFile($filename, $localname = null)
    {
        $zip = $this->zip;

        $realpathFilename = realpath($filename);
        if ($realpathFilename !== false) {
            $filename = $realpathFilename;
        }

        $filenameParts = pathinfo($filename);
        $localnameParts = pathinfo($localname);

        
        $tempFile = false;
        if ($filenameParts['basename'] != $localnameParts['basename']) {
            $tempFile = true; 
            $temppath = $this->tempDir . DIRECTORY_SEPARATOR . $localnameParts['basename'];
            copy($filename, $temppath);
            $filename = $temppath;
            $filenameParts = pathinfo($temppath);
        }

        $pathRemoved = $filenameParts['dirname'];
        $pathAdded = $localnameParts['dirname'];

        $res = $zip->add($filename, PCLZIP_OPT_REMOVE_PATH, $pathRemoved, PCLZIP_OPT_ADD_PATH, $pathAdded);

        if ($tempFile) {
            unlink($this->tempDir . DIRECTORY_SEPARATOR . $localnameParts['basename']);
        }

        return ($res == 0) ? false : true;
    }

 
    public function pclzipAddFromString($localname, $contents)
    {
        $zip = $this->zip;
        $filenameParts = pathinfo($localname);

        $handle = fopen($this->tempDir . DIRECTORY_SEPARATOR . $filenameParts['basename'], 'wb');
        fwrite($handle, $contents);
        fclose($handle);

        $filename = $this->tempDir . DIRECTORY_SEPARATOR . $filenameParts['basename'];
        $pathRemoved = $this->tempDir;
        $pathAdded = $filenameParts['dirname'];

        $res = $zip->add($filename, PCLZIP_OPT_REMOVE_PATH, $pathRemoved, PCLZIP_OPT_ADD_PATH, $pathAdded);

        @unlink($this->tempDir . DIRECTORY_SEPARATOR . $filenameParts['basename']);

        return ($res == 0) ? false : true;
    }


    public function pclzipExtractTo($destination, $entries = null)
    {
        $zip = $this->zip;

        if (is_null($entries)) {
            $result = $zip->extract(PCLZIP_OPT_PATH, $destination);
            return ($result > 0) ? true : false;
        }

        if (!is_array($entries)) {
            $entries = array($entries);
        }
        foreach ($entries as $entry) {
            $entryIndex = $this->locateName($entry);
            $result = $zip->extractByIndex($entryIndex, PCLZIP_OPT_PATH, $destination);
            if ($result <= 0) {
                return false;
            }
        }

        return true;
    }

   
    public function pclzipGetFromName($filename)
    {
        $zip = $this->zip;
        $listIndex = $this->pclzipLocateName($filename);
        $contents = false;

        if ($listIndex !== false) {
            $extracted = $zip->extractByIndex($listIndex, PCLZIP_OPT_EXTRACT_AS_STRING);
        } else {
            $filename = substr($filename, 1);
            $listIndex = $this->pclzipLocateName($filename);
            $extracted = $zip->extractByIndex($listIndex, PCLZIP_OPT_EXTRACT_AS_STRING);
        }
        if ((is_array($extracted)) && ($extracted != 0)) {
            $contents = $extracted[0]['content'];
        }

        return $contents;
    }

    
    public function pclzipGetNameIndex($index)
    {
        $zip = $this->zip;
        $list = $zip->listContent();
        if (isset($list[$index])) {
            return $list[$index]['filename'];
        } else {
            return false;
        }
    }

    
    public function pclzipLocateName($filename)
    {
        $zip = $this->zip;
        $list = $zip->listContent();
        $listCount = count($list);
        $listIndex = -1;
        for ($i = 0; $i < $listCount; ++$i) {
            if (strtolower($list[$i]['filename']) == strtolower($filename) ||
                strtolower($list[$i]['stored_filename']) == strtolower($filename)) {
                $listIndex = $i;
                break;
            }
        }

        return ($listIndex > -1) ? $listIndex : false;
    }
}
