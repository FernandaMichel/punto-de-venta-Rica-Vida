<?php


namespace PhpOffice\PhpWord\Reader;

use PhpOffice\PhpWord\Exception\Exception;

abstract class AbstractReader implements ReaderInterface
{
    
    protected $readDataOnly = true;

   
    protected $fileHandle;

   
    public function isReadDataOnly()
    {
        return true;
    }

   
    public function setReadDataOnly($value = true)
    {
        $this->readDataOnly = $value;
        return $this;
    }

    
    protected function openFile($filename)
    {
        if (!file_exists($filename) || !is_readable($filename)) {
            throw new Exception("Could not open " . $filename . " for reading! File does not exist.");
        }

        $this->fileHandle = fopen($filename, 'r');
        if ($this->fileHandle === false) {
            throw new Exception("Could not open file " . $filename . " for reading.");
        }
    }

   
    public function canRead($filename)
    {
        try {
            $this->openFile($filename);
        } catch (Exception $e) {
            return false;
        }
        if (is_resource($this->fileHandle)) {
            fclose($this->fileHandle);
        }

        return true;
    }

   
    public function getReadDataOnly()
    {
        return $this->isReadDataOnly();
    }
}
