<?php


namespace PhpOffice\PhpWord\Style;


class Chart extends AbstractStyle
{

   
    private $width = 1000000;

   
    private $height = 1000000;

    
    private $is3d = false;

   
    public function __construct($style = array())
    {
        $this->setStyleByArray($style);
    }

    
    public function getWidth()
    {
        return $this->width;
    }

    
    public function setWidth($value = null)
    {
        $this->width = $this->setIntVal($value, $this->width);

        return $this;
    }

    
    public function getHeight()
    {
        return $this->height;
    }

    
    public function setHeight($value = null)
    {
        $this->height = $this->setIntVal($value, $this->height);

        return $this;
    }

   
    public function is3d()
    {
        return $this->is3d;
    }

  
    public function set3d($value = true)
    {
        $this->is3d = $this->setBoolVal($value, $this->is3d);

        return $this;
    }
}
