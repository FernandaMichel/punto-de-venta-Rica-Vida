<?php


namespace PhpOffice\PhpWord\Metadata;

class Protection
{
  
    private $editing;

    
    public function __construct($editing = null)
    {
        $this->setEditing($editing);
    }

    
    public function getEditing()
    {
        return $this->editing;
    }

  
    public function setEditing($editing = null)
    {
        $this->editing = $editing;

        return $this;
    }
}
