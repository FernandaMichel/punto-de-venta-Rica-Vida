<?php

namespace PhpOffice\PhpWord\Style;


class Image extends Frame
{
    const WRAPPING_STYLE_INLINE = self::WRAP_INLINE;
    const WRAPPING_STYLE_SQUARE = self::WRAP_SQUARE;
    const WRAPPING_STYLE_TIGHT = self::WRAP_TIGHT;
    const WRAPPING_STYLE_BEHIND = self::WRAP_BEHIND;
    const WRAPPING_STYLE_INFRONT = self::WRAP_INFRONT;
    const POSITION_HORIZONTAL_LEFT = self::POS_LEFT;
    const POSITION_HORIZONTAL_CENTER = self::POS_CENTER;
    const POSITION_HORIZONTAL_RIGHT = self::POS_RIGHT;
    const POSITION_VERTICAL_TOP = self::POS_TOP;
    const POSITION_VERTICAL_CENTER = self::POS_CENTER;
    const POSITION_VERTICAL_BOTTOM = self::POS_BOTTOM;
    const POSITION_VERTICAL_INSIDE = self::POS_INSIDE;
    const POSITION_VERTICAL_OUTSIDE = self::POS_OUTSIDE;
    const POSITION_RELATIVE_TO_MARGIN = self::POS_RELTO_MARGIN;
    const POSITION_RELATIVE_TO_PAGE = self::POS_RELTO_PAGE;
    const POSITION_RELATIVE_TO_COLUMN = self::POS_RELTO_COLUMN;
    const POSITION_RELATIVE_TO_CHAR = self::POS_RELTO_CHAR;
    const POSITION_RELATIVE_TO_TEXT = self::POS_RELTO_TEXT;
    const POSITION_RELATIVE_TO_LINE = self::POS_RELTO_LINE;
    const POSITION_RELATIVE_TO_LMARGIN = self::POS_RELTO_LMARGIN;
    const POSITION_RELATIVE_TO_RMARGIN = self::POS_RELTO_RMARGIN;
    const POSITION_RELATIVE_TO_TMARGIN = self::POS_RELTO_TMARGIN;
    const POSITION_RELATIVE_TO_BMARGIN = self::POS_RELTO_BMARGIN;
    const POSITION_RELATIVE_TO_IMARGIN = self::POS_RELTO_IMARGIN;
    const POSITION_RELATIVE_TO_OMARGIN = self::POS_RELTO_OMARGIN;
    const POSITION_ABSOLUTE = self::POS_ABSOLUTE;
    const POSITION_RELATIVE = self::POS_RELATIVE;

   
    public function __construct()
    {
        parent::__construct();
        $this->setUnit('px');

       
        $this->setWrap(self::WRAPPING_STYLE_INLINE);
        $this->setHPos(self::POSITION_HORIZONTAL_LEFT);
        $this->setHPosRelTo(self::POSITION_RELATIVE_TO_CHAR);
        $this->setVPos(self::POSITION_VERTICAL_TOP);
        $this->setVPosRelTo(self::POSITION_RELATIVE_TO_LINE);
    }

   
    public function getMarginTop()
    {
        return $this->getTop();
    }

   
    public function setMarginTop($value = 0)
    {
        $this->setTop($value);

        return $this;
    }

    
    public function getMarginLeft()
    {
        return $this->getLeft();
    }

    
    public function setMarginLeft($value = 0)
    {
        $this->setLeft($value);

        return $this;
    }

    
    public function getWrappingStyle()
    {
        return $this->getWrap();
    }

   
    public function setWrappingStyle($wrappingStyle)
    {
        $this->setWrap($wrappingStyle);

        return $this;
    }

    
    public function getPositioning()
    {
        return $this->getPos();
    }

    
    public function setPositioning($positioning)
    {
        $this->setPos($positioning);

        return $this;
    }

    public function getPosHorizontal()
    {
        return $this->getHPos();
    }

   
    public function setPosHorizontal($alignment)
    {
        $this->setHPos($alignment);

        return $this;
    }

   
    public function getPosVertical()
    {
        return $this->getVPos();
    }

    
    public function setPosVertical($alignment)
    {
        $this->setVPos($alignment);

        return $this;
    }

   
    public function getPosHorizontalRel()
    {
        return $this->getHPosRelTo();
    }

   
    public function setPosHorizontalRel($relto)
    {
        $this->setHPosRelTo($relto);

        return $this;
    }

   
    public function getPosVerticalRel()
    {
        return $this->getVPosRelTo();
    }

  
    public function setPosVerticalRel($relto)
    {
        $this->setVPosRelTo($relto);

        return $this;
    }
}
