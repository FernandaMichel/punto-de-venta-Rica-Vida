<?php


namespace PhpOffice\PhpWord\Style;


class Extrusion extends AbstractStyle
{
    
    const EXTRUSION_PARALLEL = 'parallel';
    const EXTRUSION_PERSPECTIVE = 'perspective';

    
    private $type;

    
    private $color;

    
    public function __construct($style = array())
    {
        $this->setStyleByArray($style);
    }

    
    public function getType()
    {
        return $this->type;
    }

    
    public function setType($value = null)
    {
        $enum = array(self::EXTRUSION_PARALLEL, self::EXTRUSION_PERSPECTIVE);
        $this->type = $this->setEnumVal($value, $enum, null);

        return $this;
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
