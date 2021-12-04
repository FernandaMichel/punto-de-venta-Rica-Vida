<?php

namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Style\Line as LineStyle;


class Line extends AbstractElement
{
    
    private $style;

    
    public function __construct($style = null)
    {
        $this->style = $this->setNewStyle(new LineStyle(), $style);
    }

    
    public function getStyle()
    {
        return $this->style;
    }
}
