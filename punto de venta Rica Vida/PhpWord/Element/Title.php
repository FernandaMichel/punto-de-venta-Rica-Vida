<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Shared\String;
use PhpOffice\PhpWord\Style;


class Title extends AbstractElement
{
    
    private $text;

    
    private $depth = 1;

    
    private $style;

    
    protected $collectionRelation = true;

    
    public function __construct($text, $depth = 1)
    {
        $this->text = String::toUTF8($text);
        $this->depth = $depth;
        if (array_key_exists("Heading_{$this->depth}", Style::getStyles())) {
            $this->style = "Heading{$this->depth}";
        }

        return $this;
    }

   
    public function getText()
    {
        return $this->text;
    }

    
    public function getDepth()
    {
        return $this->depth;
    }

    public function getStyle()
    {
        return $this->style;
    }
}
