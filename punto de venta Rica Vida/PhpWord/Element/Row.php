<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Style\Row as RowStyle;


class Row extends AbstractElement
{
    
    private $height = null;

   
    private $style;

    
    private $cells = array();

    
    public function __construct($height = null, $style = null)
    {
        $this->height = $height;
        $this->style = $this->setNewStyle(new RowStyle(), $style, true);
    }

    public function addCell($width = null, $style = null)
    {
        $cell = new Cell($width, $style);
        $cell->setParentContainer($this);
        $this->cells[] = $cell;

        return $cell;
    }

    
    public function getCells()
    {
        return $this->cells;
    }

    
    public function getStyle()
    {
        return $this->style;
    }

   
    public function getHeight()
    {
        return $this->height;
    }
}
