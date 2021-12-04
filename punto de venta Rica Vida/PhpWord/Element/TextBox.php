<?php

namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Style\TextBox as TextBoxStyle;


class TextBox extends AbstractContainer
{
    
    protected $container = 'TextBox';

    
    private $style;

    
    public function __construct($style = null)
    {
        $this->style = $this->setNewStyle(new TextBoxStyle(), $style);
    }

    
    public function getStyle()
    {
        return $this->style;
    }
}
