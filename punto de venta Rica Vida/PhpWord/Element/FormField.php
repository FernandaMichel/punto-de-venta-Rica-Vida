<?php


namespace PhpOffice\PhpWord\Element;


class FormField extends Text
{
    
    private $type = 'textinput';

    
    private $name;

    
    private $default;

   
    private $value;

    
    private $entries = array();

    
    public function __construct($type, $fontStyle = null, $paragraphStyle = null)
    {
        $this->setType($type);
    }

    
    public function getType()
    {
        return $this->type;
    }

    
    public function setType($value)
    {
        $enum = array('textinput', 'checkbox', 'dropdown');
        $this->type = $this->setEnumVal($value, $enum, $this->type);

        return $this;
    }

   
    public function getName()
    {
        return $this->name;
    }

    public function setName($value)
    {
        $this->name = $value;

        return $this;
    }

    
    public function getDefault()
    {
        return $this->default;
    }

    
    public function setDefault($value)
    {
        $this->default = $value;

        return $this;
    }

   
    public function getValue()
    {
        return $this->value;
    }

    
    public function setValue($value)
    {
        $this->value = $value;

        return $this;
    }

    
    public function getEntries()
    {
        return $this->entries;
    }

    
    public function setEntries($value)
    {
        $this->entries = $value;

        return $this;
    }
}
