<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Style\Table as TableStyle;


class Table extends AbstractElement
{
    
    private $style;

   
    private $rows = array();

    
    private $width = null;

    
    public function __construct($style = null)
    {
        $this->style = $this->setNewStyle(new TableStyle(), $style);
    }

    
    public function addRow($height = null, $style = null)
    {
        $row = new Row($height, $style);
        $row->setParentContainer($this);
        $this->rows[] = $row;

        return $row;
    }

    
    public function addCell($width = null, $style = null)
    {
        $index = count($this->rows) - 1;
        $row = $this->rows[$index];
        $cell = $row->addCell($width, $style);

        return $cell;
    }

    
    public function getRows()
    {
        return $this->rows;
    }

   
    public function getStyle()
    {
        return $this->style;
    }

   
    public function getWidth()
    {
        return $this->width;
    }

    
    public function setWidth($width)
    {
        $this->width = $width;
    }

   
    public function countColumns()
    {
        $columnCount = 0;
        if (is_array($this->rows)) {
            $rowCount = count($this->rows);
            for ($i = 0; $i < $rowCount; $i++) {
                $row = $this->rows[$i];
                $cellCount = count($row->getCells());
                if ($columnCount < $cellCount) {
                    $columnCount = $cellCount;
                }
            }
        }

        return $columnCount;
    }
}
