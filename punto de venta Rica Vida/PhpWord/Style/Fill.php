<?php


namespace PhpOffice\PhpWord\Style;


class Fill extends AbstractStyle
{
    
    private $color;

    
    public function __construct($style = array())
    {
        $this->setStyleByArray($style);
    }

    
    public function getColor()
    {
        return $this->color;
    }

    public function setColor($value = null)
    {
        $this->color = $value;

        return $this;
    }
}
