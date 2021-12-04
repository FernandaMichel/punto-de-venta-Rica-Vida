<?php


namespace PhpOffice\PhpWord\Style;


class Border extends AbstractStyle
{
    
    protected $borderTopSize;

    
    protected $borderTopColor;

   
    protected $borderLeftSize;

    
    protected $borderLeftColor;

    
    protected $borderRightSize;

    
    protected $borderRightColor;

    
    protected $borderBottomSize;

   
    protected $borderBottomColor;

   
    public function getBorderSize()
    {
        return array(
            $this->getBorderTopSize(),
            $this->getBorderLeftSize(),
            $this->getBorderRightSize(),
            $this->getBorderBottomSize(),
        );
    }

    
    public function setBorderSize($value = null)
    {
        $this->setBorderTopSize($value);
        $this->setBorderLeftSize($value);
        $this->setBorderRightSize($value);
        $this->setBorderBottomSize($value);

        return $this;
    }

    
    public function getBorderColor()
    {
        return array(
            $this->getBorderTopColor(),
            $this->getBorderLeftColor(),
            $this->getBorderRightColor(),
            $this->getBorderBottomColor(),
        );
    }

    
    public function setBorderColor($value = null)
    {
        $this->setBorderTopColor($value);
        $this->setBorderLeftColor($value);
        $this->setBorderRightColor($value);
        $this->setBorderBottomColor($value);

        return $this;
    }

    
    public function getBorderTopSize()
    {
        return $this->borderTopSize;
    }

   
    public function setBorderTopSize($value = null)
    {
        $this->borderTopSize = $this->setNumericVal($value, $this->borderTopSize);

        return $this;
    }

  
    public function getBorderTopColor()
    {
        return $this->borderTopColor;
    }

    
    public function setBorderTopColor($value = null)
    {
        $this->borderTopColor = $value;

        return $this;
    }

    public function getBorderLeftSize()
    {
        return $this->borderLeftSize;
    }

    
    public function setBorderLeftSize($value = null)
    {
        $this->borderLeftSize = $this->setNumericVal($value, $this->borderLeftSize);

        return $this;
    }

    
    public function getBorderLeftColor()
    {
        return $this->borderLeftColor;
    }

   
    public function setBorderLeftColor($value = null)
    {
        $this->borderLeftColor = $value;

        return $this;
    }

   
    public function getBorderRightSize()
    {
        return $this->borderRightSize;
    }

    
    public function setBorderRightSize($value = null)
    {
        $this->borderRightSize = $this->setNumericVal($value, $this->borderRightSize);

        return $this;
    }

    
    public function getBorderRightColor()
    {
        return $this->borderRightColor;
    }

    public function setBorderRightColor($value = null)
    {
        $this->borderRightColor = $value;

        return $this;
    }

    
    public function getBorderBottomSize()
    {
        return $this->borderBottomSize;
    }

    
    public function setBorderBottomSize($value = null)
    {
        $this->borderBottomSize = $this->setNumericVal($value, $this->borderBottomSize);

        return $this;
    }

   
    public function getBorderBottomColor()
    {
        return $this->borderBottomColor;
    }

    
    public function setBorderBottomColor($value = null)
    {
        $this->borderBottomColor = $value;

        return $this;
    }

   
    public function hasBorder()
    {
        $borders = $this->getBorderSize();

        return $borders !== array_filter($borders, 'is_null');
    }
}
