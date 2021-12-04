<?php


namespace PhpOffice\PhpWord\Style;


class Frame extends AbstractStyle
{
    
    const UNIT_PT = 'pt'; 
    const UNIT_PX = 'px'; 

    
    const POS_ABSOLUTE = 'absolute';
    const POS_RELATIVE = 'relative';

    
    const POS_CENTER = 'center';
    const POS_LEFT = 'left';
    const POS_RIGHT = 'right';
    const POS_TOP = 'top';
    const POS_BOTTOM = 'bottom';
    const POS_INSIDE = 'inside';
    const POS_OUTSIDE = 'outside';

   
    const POS_RELTO_MARGIN = 'margin';
    const POS_RELTO_PAGE = 'page';
    const POS_RELTO_COLUMN = 'column'; 
    const POS_RELTO_CHAR = 'char'; 
    const POS_RELTO_TEXT = 'text'; 
    const POS_RELTO_LINE = 'line'; 
    const POS_RELTO_LMARGIN = 'left-margin-area'; 
    const POS_RELTO_RMARGIN = 'right-margin-area'; 
    const POS_RELTO_TMARGIN = 'top-margin-area'; 
    const POS_RELTO_BMARGIN = 'bottom-margin-area'; 
    const POS_RELTO_IMARGIN = 'inner-margin-area';
    const POS_RELTO_OMARGIN = 'outer-margin-area';

    
    const WRAP_INLINE = 'inline';
    const WRAP_SQUARE = 'square';
    const WRAP_TIGHT = 'tight';
    const WRAP_THROUGH = 'through';
    const WRAP_TOPBOTTOM = 'topAndBottom';
    const WRAP_BEHIND = 'behind';
    const WRAP_INFRONT = 'infront';

    
    private $alignment;

    
    private $unit = 'pt';

   
    private $width;

    
    private $height;

    
    private $left = 0;

   
    private $top = 0;

    
    private $pos;

    
    private $hPos;

    
    private $hPosRelTo;

    
    private $vPos;

   
    private $vPosRelTo;

    
    private $wrap;

    
    public function __construct($style = array())
    {
        $this->alignment = new Alignment();
        $this->setStyleByArray($style);
    }

    
    public function getAlign()
    {
        return $this->alignment->getValue();
    }

    
    public function setAlign($value = null)
    {
        $this->alignment->setValue($value);

        return $this;
    }

   
    public function getUnit()
    {
        return $this->unit;
    }

    
    public function setUnit($value)
    {
        $this->unit = $value;

        return $this;
    }

   
    public function getWidth()
    {
        return $this->width;
    }

    
    public function setWidth($value = null)
    {
        $this->width = $this->setNumericVal($value, null);

        return $this;
    }

    
    public function getHeight()
    {
        return $this->height;
    }

    
    public function setHeight($value = null)
    {
        $this->height = $this->setNumericVal($value, null);

        return $this;
    }

    
    public function getLeft()
    {
        return $this->left;
    }

   
    public function setLeft($value = 0)
    {
        $this->left = $this->setNumericVal($value, 0);

        return $this;
    }

   
    public function getTop()
    {
        return $this->top;
    }

   
    public function setTop($value = 0)
    {
        $this->top = $this->setNumericVal($value, 0);

        return $this;
    }

   
    public function getPos()
    {
        return $this->pos;
    }

    
    public function setPos($value)
    {
        $enum = array(
            self::POS_ABSOLUTE,
            self::POS_RELATIVE,
        );
        $this->pos = $this->setEnumVal($value, $enum, $this->pos);

        return $this;
    }

    
    public function getHPos()
    {
        return $this->hPos;
    }

    public function setHPos($value)
    {
        $enum = array(
            self::POS_ABSOLUTE,
            self::POS_LEFT,
            self::POS_CENTER,
            self::POS_RIGHT,
            self::POS_INSIDE,
            self::POS_OUTSIDE,
        );
        $this->hPos = $this->setEnumVal($value, $enum, $this->hPos);

        return $this;
    }

   
    public function getVPos()
    {
        return $this->vPos;
    }

    
    public function setVPos($value)
    {
        $enum = array(
            self::POS_ABSOLUTE,
            self::POS_TOP,
            self::POS_CENTER,
            self::POS_BOTTOM,
            self::POS_INSIDE,
            self::POS_OUTSIDE,
        );
        $this->vPos = $this->setEnumVal($value, $enum, $this->vPos);

        return $this;
    }

   
    public function getHPosRelTo()
    {
        return $this->hPosRelTo;
    }

    
    public function setHPosRelTo($value)
    {
        $enum = array(
            self::POS_RELTO_MARGIN,
            self::POS_RELTO_PAGE,
            self::POS_RELTO_COLUMN,
            self::POS_RELTO_CHAR,
            self::POS_RELTO_LMARGIN,
            self::POS_RELTO_RMARGIN,
            self::POS_RELTO_IMARGIN,
            self::POS_RELTO_OMARGIN,
        );
        $this->hPosRelTo = $this->setEnumVal($value, $enum, $this->hPosRelTo);

        return $this;
    }

   
    public function getVPosRelTo()
    {
        return $this->vPosRelTo;
    }

    public function setVPosRelTo($value)
    {
        $enum = array(
            self::POS_RELTO_MARGIN,
            self::POS_RELTO_PAGE,
            self::POS_RELTO_TEXT,
            self::POS_RELTO_LINE,
            self::POS_RELTO_TMARGIN,
            self::POS_RELTO_BMARGIN,
            self::POS_RELTO_IMARGIN,
            self::POS_RELTO_OMARGIN,
        );
        $this->vPosRelTo = $this->setEnumVal($value, $enum, $this->vPosRelTo);

        return $this;
    }

  
    public function getWrap()
    {
        return $this->wrap;
    }

   
    public function setWrap($value)
    {
        $enum = array(
            self::WRAP_INLINE,
            self::WRAP_SQUARE,
            self::WRAP_TIGHT,
            self::WRAP_THROUGH,
            self::WRAP_TOPBOTTOM,
            self::WRAP_BEHIND,
            self::WRAP_INFRONT,
        );
        $this->wrap = $this->setEnumVal($value, $enum, $this->wrap);

        return $this;
    }
}
