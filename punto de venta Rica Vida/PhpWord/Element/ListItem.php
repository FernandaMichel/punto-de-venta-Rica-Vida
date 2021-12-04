<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Shared\String;
use PhpOffice\PhpWord\Style\ListItem as ListItemStyle;


class ListItem extends AbstractElement
{
    
    private $style;

    private $textObject;

    
    private $depth;

    public function __construct($text, $depth = 0, $fontStyle = null, $listStyle = null, $paragraphStyle = null)
    {
        $this->textObject = new Text(String::toUTF8($text), $fontStyle, $paragraphStyle);
        $this->depth = $depth;

        if (!is_null($listStyle) && is_string($listStyle)) {
            $this->style = new ListItemStyle($listStyle);
        } else {
            $this->style = $this->setNewStyle(new ListItemStyle(), $listStyle, true);
        }
    }

    
    public function getStyle()
    {
        return $this->style;
    }


    public function getTextObject()
    {
        return $this->textObject;
    }

    
    public function getDepth()
    {
        return $this->depth;
    }

    public function getText()
    {
        return $this->textObject->getText();
    }
}
