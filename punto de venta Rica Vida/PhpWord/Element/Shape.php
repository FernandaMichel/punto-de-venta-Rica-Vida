<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Style\Shape as ShapeStyle;


class Shape extends AbstractElement
{
    
    private $type;

    
    private $style;

    
    public function __construct($type, $style = null)
    {
        $this->setType($type);
        $this->style = $this->setNewStyle(new ShapeStyle(), $style);
    }

    
    public function getType()
    {
        return $this->type;
    }

    
    public function setType($value = null)
    {
        $enum = array('arc', 'curve', 'line', 'polyline', 'rect', 'oval');
        $this->type = $this->setEnumVal($value, $enum, null);

        return $this;
    }

    public function getStyle()
    {
        return $this->style;
    }
}
