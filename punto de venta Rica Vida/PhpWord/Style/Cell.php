<?php


namespace PhpOffice\PhpWord\Style;


class Cell extends Border
{
    
    const VALIGN_TOP = 'top';
    const VALIGN_CENTER = 'center';
    const VALIGN_BOTTOM = 'bottom';
    const VALIGN_BOTH = 'both';

    
    const TEXT_DIR_BTLR = 'btLr';
    const TEXT_DIR_TBRL = 'tbRl';

   
    const VMERGE_RESTART = 'restart';
    const VMERGE_CONTINUE = 'continue';

    
    const DEFAULT_BORDER_COLOR = '000000';

    
    private $vAlign;

    
    private $textDirection;

    
    private $gridSpan;

    
    private $vMerge;

   
    private $shading;

    
    public function getVAlign()
    {
        return $this->vAlign;
    }

    
    public function setVAlign($value = null)
    {
        $enum = array(self::VALIGN_TOP, self::VALIGN_CENTER, self::VALIGN_BOTTOM, self::VALIGN_BOTH);
        $this->vAlign = $this->setEnumVal($value, $enum, $this->vAlign);

        return $this;
    }

    
    public function getTextDirection()
    {
        return $this->textDirection;
    }

    
    public function setTextDirection($value = null)
    {
        $enum = array(self::TEXT_DIR_BTLR, self::TEXT_DIR_TBRL);
        $this->textDirection = $this->setEnumVal($value, $enum, $this->textDirection);

        return $this;
    }

   
    public function getBgColor()
    {
        if ($this->shading !== null) {
            return $this->shading->getFill();
        } else {
            return null;
        }
    }

    
    public function setBgColor($value = null)
    {
        return $this->setShading(array('fill' => $value));
    }

    
    public function getGridSpan()
    {
        return $this->gridSpan;
    }

    
    public function setGridSpan($value = null)
    {
        $this->gridSpan = $this->setIntVal($value, $this->gridSpan);

        return $this;
    }

    
    public function getVMerge()
    {
        return $this->vMerge;
    }

    
    public function setVMerge($value = null)
    {
        $enum = array(self::VMERGE_RESTART, self::VMERGE_CONTINUE);
        $this->vMerge = $this->setEnumVal($value, $enum, $this->vMerge);

        return $this;
    }

   
    public function getShading()
    {
        return $this->shading;
    }

    
    public function setShading($value = null)
    {
        $this->setObjectVal($value, 'Shading', $this->shading);

        return $this;
    }

    
    public function getDefaultBorderColor()
    {
        return self::DEFAULT_BORDER_COLOR;
    }
}
