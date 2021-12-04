<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Shared\String;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Style\Paragraph;


class Link extends AbstractElement
{
    
    private $source;

   
    private $text;

    
    private $fontStyle;

    
    private $paragraphStyle;

   
    protected $mediaRelation = true;

    
    protected $internal = false;

    
    public function __construct($source, $text = null, $fontStyle = null, $paragraphStyle = null, $internal = false)
    {
        $this->source = String::toUTF8($source);
        $this->text = is_null($text) ? $this->source : String::toUTF8($text);
        $this->fontStyle = $this->setNewStyle(new Font('text'), $fontStyle);
        $this->paragraphStyle = $this->setNewStyle(new Paragraph(), $paragraphStyle);
        $this->internal = $internal;
        return $this;
    }

    
    public function getSource()
    {
        return $this->source;
    }

    
    public function getText()
    {
        return $this->text;
    }

    
    public function getFontStyle()
    {
        return $this->fontStyle;
    }

    
    public function getParagraphStyle()
    {
        return $this->paragraphStyle;
    }

    
    public function getTarget()
    {
        return $this->source;
    }

    
    public function getLinkSrc()
    {
        return $this->getSource();
    }

    
    public function getLinkName()
    {
        return $this->getText();
    }

   
    public function isInternal()
    {
        return $this->internal;
    }
}
