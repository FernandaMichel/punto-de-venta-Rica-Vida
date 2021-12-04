<?php


namespace PhpOffice\PhpWord\Element;


class SDT extends Text
{
    
    private $type;

   
    private $value;

   
    private $listItems = array();

    
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
        $enum = array('comboBox', 'dropDownList', 'date');
        $this->type = $this->setEnumVal($value, $enum, 'comboBox');

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

    
    public function getListItems()
    {
        return $this->listItems;
    }

    
    public function setListItems($value)
    {
        $this->listItems = $value;

        return $this;
    }
}
