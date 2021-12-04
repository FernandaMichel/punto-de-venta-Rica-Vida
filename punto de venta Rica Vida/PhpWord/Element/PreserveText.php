<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Shared\String;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Style\Paragraph;


class PreserveText extends AbstractElement
{
    
    private $text;

    
    private $fontStyle;

    
    private $paragraphStyle;

    public function __construct($text = null, $fontStyle = null, $paragraphStyle = null)
    {
        $this->fontStyle = $this->setNewStyle(new Font('text'), $fontStyle);
        $this->paragraphStyle = $this->setNewStyle(new Paragraph(), $paragraphStyle);

        $this->text = String::toUTF8($text);
        $matches = preg_split('/({.*?})/', $this->text, null, PREG_SPLIT_DELIM_CAPTURE | PREG_SPLIT_NO_EMPTY);
        if (isset($matches[0])) {
            $this->text = $matches;
        }

        return $this;
    }

   
    public function getFontStyle()
    {
        return $this->fontStyle;
    }

    
    public function getParagraphStyle()
    {
        return $this->paragraphStyle;
    }

    
    public function getText()
    {
        return $this->text;
    }
}
