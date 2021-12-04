<?php


namespace PhpOffice\PhpWord\Metadata;


class Compatibility
{
    
    private $ooxmlVersion = 12;

    
    public function getOoxmlVersion()
    {
        return $this->ooxmlVersion;
    }

    
    public function setOoxmlVersion($value)
    {
        $this->ooxmlVersion = $value;

        return $this;
    }
}
