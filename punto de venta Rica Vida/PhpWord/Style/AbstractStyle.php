<?php


namespace PhpOffice\PhpWord\Style;

use PhpOffice\PhpWord\Shared\String;


abstract class AbstractStyle
{
    
    protected $styleName;

    
    protected $index;

   
    protected $aliases = array();

    
    private $isAuto = false;

    
    public function getStyleName()
    {
        return $this->styleName;
    }

    
    public function setStyleName($value)
    {
        $this->styleName = $value;

        return $this;
    }

    
    public function getIndex()
    {
        return $this->index;
    }
    public function setIndex($value = null)
    {
        $this->index = $this->setIntVal($value, $this->index);

        return $this;
    }

   
    public function isAuto()
    {
        return $this->isAuto;
    }

    
    public function setAuto($value = true)
    {
        $this->isAuto = $this->setBoolVal($value, $this->isAuto);

        return $this;
    }

    
    public function getChildStyleValue($substyleObject, $substyleProperty)
    {
        if ($substyleObject !== null) {
            $method = "get{$substyleProperty}";
            return $substyleObject->$method();
        } else {
            return null;
        }
    }

   
    public function setStyleValue($key, $value)
    {
        if (isset($this->aliases[$key])) {
            $key = $this->aliases[$key];
        }
        $method = 'set' . String::removeUnderscorePrefix($key);
        if (method_exists($this, $method)) {
            $this->$method($value);
        }

        return $this;
    }

    
    public function setStyleByArray($values = array())
    {
        foreach ($values as $key => $value) {
            $this->setStyleValue($key, $value);
        }

        return $this;
    }

    
    protected function setNonEmptyVal($value, $default)
    {
        if ($value === null || $value == '') {
            $value = $default;
        }

        return $value;
    }

    
    protected function setBoolVal($value, $default)
    {
        if (!is_bool($value)) {
            $value = $default;
        }

        return $value;
    }

    
    protected function setNumericVal($value, $default = null)
    {
        if (!is_numeric($value)) {
            $value = $default;
        }

        return $value;
    }

    
    protected function setIntVal($value, $default = null)
    {
        if (is_string($value) && (preg_match('/[^\d]/', $value) == 0)) {
            $value = intval($value);
        }
        if (!is_numeric($value)) {
            $value = $default;
        } else {
            $value = intval($value);
        }

        return $value;
    }

   
    protected function setFloatVal($value, $default = null)
    {
        if (is_string($value) && (preg_match('/[^\d\.\,]/', $value) == 0)) {
            $value = floatval($value);
        }
        if (!is_numeric($value)) {
            $value = $default;
        }

        return $value;
    }

    
    protected function setEnumVal($value = null, $enum = array(), $default = null)
    {
        if ($value != null && trim($value) != '' && !empty($enum) && !in_array($value, $enum)) {
            throw new \InvalidArgumentException("Invalid style value: {$value} Options:".join(',', $enum));
        } elseif ($value === null || trim($value) == '') {
            $value = $default;
        }

        return $value;
    }

   
    protected function setObjectVal($value, $styleName, &$style)
    {
        $styleClass = substr(get_class($this), 0, strrpos(get_class($this), '\\')) . '\\' . $styleName;
        if (is_array($value)) {
            if (!$style instanceof $styleClass) {
                $style = new $styleClass();
            }
            $style->setStyleByArray($value);
        } else {
            $style = $value;
        }

        return $style;
    }

    
    protected function setPairedVal(&$property, &$pairProperty, $value)
    {
        $property = $this->setBoolVal($value, $property);
        if ($value == true) {
            $pairProperty = false;
        }

        return $this;
    }

    public function setArrayStyle(array $style = array())
    {
        return $this->setStyleByArray($style);
    }
}
