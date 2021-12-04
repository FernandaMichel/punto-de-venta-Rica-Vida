<?php


namespace PhpOffice\PhpWord\Style;


class LineNumbering extends AbstractStyle
{
    const LINE_NUMBERING_CONTINUOUS  = 'continuous';
    const LINE_NUMBERING_NEW_PAGE    = 'newPage';
    const LINE_NUMBERING_NEW_SECTION = 'newSection';

    private $start = 1;

   
    private $increment = 1;

    
    private $distance;

    
    private $restart;

    
    public function __construct($style = array())
    {
        $this->setStyleByArray($style);
    }

    
    public function getStart()
    {
        return $this->start;
    }

    
    public function setStart($value = null)
    {
        $this->start = $this->setIntVal($value, $this->start);

        return $this;
    }

   
    public function getIncrement()
    {
        return $this->increment;
    }

   
    public function setIncrement($value = null)
    {
        $this->increment = $this->setIntVal($value, $this->increment);

        return $this;
    }

   
    public function getDistance()
    {
        return $this->distance;
    }

   
    public function setDistance($value = null)
    {
        $this->distance = $this->setNumericVal($value, $this->distance);

        return $this;
    }

   
    public function getRestart()
    {
        return $this->restart;
    }

    public function setRestart($value = null)
    {
        $enum = array(self::LINE_NUMBERING_CONTINUOUS, self::LINE_NUMBERING_NEW_PAGE, self::LINE_NUMBERING_NEW_SECTION);
        $this->restart = $this->setEnumVal($value, $enum, $this->restart);

        return $this;
    }
}
