<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Style\TOC as TOCStyle;


class TOC extends AbstractElement
{
    
    private $TOCStyle;

   
    private $fontStyle;

   
    private $minDepth = 1;

    private $maxDepth = 9;


  
    public function __construct($fontStyle = null, $tocStyle = null, $minDepth = 1, $maxDepth = 9)
    {
        $this->TOCStyle = new TOCStyle();

        if (!is_null($tocStyle) && is_array($tocStyle)) {
            $this->TOCStyle->setStyleByArray($tocStyle);
        }

        if (!is_null($fontStyle) && is_array($fontStyle)) {
            $this->fontStyle = new Font();
            $this->fontStyle->setStyleByArray($fontStyle);
        } else {
            $this->fontStyle = $fontStyle;
        }

        $this->minDepth = $minDepth;
        $this->maxDepth = $maxDepth;
    }

    
    public function getTitles()
    {
        if (!$this->phpWord instanceof PhpWord) {
            return array();
        }

        $titles = $this->phpWord->getTitles()->getItems();
        foreach ($titles as $i => $title) {
            $depth = $title->getDepth();
            if ($this->minDepth > $depth) {
                unset($titles[$i]);
            }
            if (($this->maxDepth != 0) && ($this->maxDepth < $depth)) {
                unset($titles[$i]);
            }
        }

        return $titles;
    }

   
    public function getStyleTOC()
    {
        return $this->TOCStyle;
    }

    
    public function getStyleFont()
    {
        return $this->fontStyle;
    }

   
    public function setMaxDepth($value)
    {
        $this->maxDepth = $value;
    }

    
    public function getMaxDepth()
    {
        return $this->maxDepth;
    }

    public function setMinDepth($value)
    {
        $this->minDepth = $value;
    }

    public function getMinDepth()
    {
        return $this->minDepth;
    }
}
