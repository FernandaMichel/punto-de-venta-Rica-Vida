<?php


namespace PhpOffice\PhpWord\Shared;

use PhpOffice\PhpWord\Settings;


class XMLWriter
{
    const STORAGE_MEMORY = 1;
    const STORAGE_DISK = 2;

   
    private $xmlWriter;

    
    private $tempFile = '';

    
    public function __construct($tempLocation = self::STORAGE_MEMORY, $tempFolder = './')
    {
        $this->xmlWriter = new \XMLWriter();

        if ($tempLocation == self::STORAGE_MEMORY) {
            $this->xmlWriter->openMemory();
        } else {
            $this->tempFile = tempnam($tempFolder, 'xml');

            
            if (false === $this->tempFile || false === $this->xmlWriter->openUri($this->tempFile)) {
                $this->xmlWriter->openMemory();
            }
        }

        $compatibility = Settings::hasCompatibility();
        if ($compatibility) {
            $this->xmlWriter->setIndent(false);
            $this->xmlWriter->setIndentString('');
        } else {
            $this->xmlWriter->setIndent(true);
            $this->xmlWriter->setIndentString('  ');
        }
    }

   
    public function __destruct()
    {
        unset($this->xmlWriter);

        if ($this->tempFile != '') {
            @unlink($this->tempFile);
        }
    }

   
    public function __call($function, $args)
    {
        if (method_exists($this->xmlWriter, $function) === false) {
            throw new \BadMethodCallException("Method '{$function}' does not exists.");
        }

        try {
            @call_user_func_array(array($this->xmlWriter, $function), $args);
        } catch (\Exception $ex) {
        }
    }

    
    public function getData()
    {
        if ($this->tempFile == '') {
            return $this->xmlWriter->outputMemory(true);
        } else {
            $this->xmlWriter->flush();
            return file_get_contents($this->tempFile);
        }
    }

    
    public function writeElementBlock($element, $attributes, $value = null)
    {
        $this->xmlWriter->startElement($element);
        if (!is_array($attributes)) {
            $attributes = array($attributes => $value);
        }
        foreach ($attributes as $attribute => $value) {
            $this->xmlWriter->writeAttribute($attribute, $value);
        }
        $this->xmlWriter->endElement();
    }

    
    public function writeElementIf($condition, $element, $attribute = null, $value = null)
    {
        if ($condition == true) {
            if (is_null($attribute)) {
                $this->xmlWriter->writeElement($element, $value);
            } else {
                $this->xmlWriter->startElement($element);
                $this->xmlWriter->writeAttribute($attribute, $value);
                $this->xmlWriter->endElement();
            }
        }
    }

    
    public function writeAttributeIf($condition, $attribute, $value)
    {
        if ($condition == true) {
            $this->xmlWriter->writeAttribute($attribute, $value);
        }
    }
}
