<?php

namespace PhpOffice\PhpWord\Style;


class Indentation extends AbstractStyle
{
    
    private $left = 0;

    
    private $right = 0;

   
    private $firstLine;

   
    private $hanging;

   
    public function __construct($style = array())
    {
        $this->setStyleByArray($style);
    }

    
    public function getLeft()
    {
        return $this->left;
    }

    
    public function setLeft($value = null)
    {
        $this->left = $this->setNumericVal($value, $this->left);

        return $this;
    }

   
    public function getRight()
    {
        return $this->right;
    }

  
    public function setRight($value = null)
    {
        $this->right = $this->setNumericVal($value, $this->right);

        return $this;
    }

    
    public function getFirstLine()
    {
        return $this->firstLine;
    }

    
    public function setFirstLine($value = null)
    {
        $this->firstLine = $this->setNumericVal($value, $this->firstLine);

        return $this;
    }

   
    public function getHanging()
    {
        return $this->hanging;
    }

    
    public function setHanging($value = null)
    {
        $this->hanging = $this->setNumericVal($value, $this->hanging);

        return $this;
    }
}
