<?php

namespace PhpOffice\PhpWord\Style;


class Line extends Image
{
   
    const CONNECTOR_TYPE_STRAIGHT = 'straight';

    
    const ARROW_STYLE_BLOCK = 'block';
    const ARROW_STYLE_OPEN = 'open';
    const ARROW_STYLE_CLASSIC = 'classic';
    const ARROW_STYLE_DIAMOND = 'diamond';
    const ARROW_STYLE_OVAL = 'oval';

  
    const DASH_STYLE_DASH = 'dash';
    const DASH_STYLE_ROUND_DOT = 'rounddot';
    const DASH_STYLE_SQUARE_DOT = 'squaredot';
    const DASH_STYLE_DASH_DOT = 'dashdot';
    const DASH_STYLE_LONG_DASH = 'longdash';
    const DASH_STYLE_LONG_DASH_DOT = 'longdashdot';
    const DASH_STYLE_LONG_DASH_DOT_DOT = 'longdashdotdot';

 
    private $flip = false;

    private $connectorType = self::CONNECTOR_TYPE_STRAIGHT;


    private $weight;

  
    private $color;

    private $dash;

    
    private $beginArrow;

    
    private $endArrow;

    
    public function isFlip()
    {
        return $this->flip;
    }

    
    public function setFlip($value = false)
    {
        $this->flip = $this->setBoolVal($value, $this->flip);

        return $this;
    }

   
    public function getConnectorType()
    {
        return $this->connectorType;
    }

    
    public function setConnectorType($value = null)
    {
        $enum = array(
            self::CONNECTOR_TYPE_STRAIGHT
        );
        $this->connectorType = $this->setEnumVal($value, $enum, $this->connectorType);

        return $this;
    }

   
    public function getWeight()
    {
        return $this->weight;
    }

   
    public function setWeight($value = null)
    {
        $this->weight = $this->setNumericVal($value, $this->weight);

        return $this;
    }

 
    public function getColor()
    {
        return $this->color;
    }

    
    public function setColor($value = null)
    {
        $this->color = $value;

        return $this;
    }

    
    public function getBeginArrow()
    {
        return $this->beginArrow;
    }

    
    public function setBeginArrow($value = null)
    {
        $enum = array(
            self::ARROW_STYLE_BLOCK, self::ARROW_STYLE_CLASSIC, self::ARROW_STYLE_DIAMOND,
            self::ARROW_STYLE_OPEN, self::ARROW_STYLE_OVAL
        );
        $this->beginArrow = $this->setEnumVal($value, $enum, $this->beginArrow);

        return $this;
    }

   
    public function getEndArrow()
    {
        return $this->endArrow;
    }

  
    public function setEndArrow($value = null)
    {
        $enum = array(
            self::ARROW_STYLE_BLOCK, self::ARROW_STYLE_CLASSIC, self::ARROW_STYLE_DIAMOND,
            self::ARROW_STYLE_OPEN, self::ARROW_STYLE_OVAL
        );
        $this->endArrow = $this->setEnumVal($value, $enum, $this->endArrow);

        return $this;
    }

    public function getDash()
    {
        return $this->dash;
    }

   
    public function setDash($value = null)
    {
        $enum = array(
            self::DASH_STYLE_DASH, self::DASH_STYLE_DASH_DOT, self::DASH_STYLE_LONG_DASH,
            self::DASH_STYLE_LONG_DASH_DOT, self::DASH_STYLE_LONG_DASH_DOT_DOT, self::DASH_STYLE_ROUND_DOT,
            self::DASH_STYLE_SQUARE_DOT
        );
        $this->dash = $this->setEnumVal($value, $enum, $this->dash);

        return $this;
    }
}
