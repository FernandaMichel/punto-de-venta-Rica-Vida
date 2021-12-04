<?php


namespace PhpOffice\PhpWord\Metadata;


class DocInfo
{
    const PROPERTY_TYPE_BOOLEAN = 'b';
    const PROPERTY_TYPE_INTEGER = 'i';
    const PROPERTY_TYPE_FLOAT = 'f';
    const PROPERTY_TYPE_DATE = 'd';
    const PROPERTY_TYPE_STRING = 's';
    const PROPERTY_TYPE_UNKNOWN = 'u';

  
    private $creator;

  
    private $lastModifiedBy;

    
    private $created;

    
    private $modified;

    
    private $title;

    
    private $description;

    private $subject;

  
    private $keywords;

    private $category;

    private $company;

    private $manager;

    
    private $customProperties = array();

    
    public function __construct()
    {
        $this->creator        = '';
        $this->lastModifiedBy = $this->creator;
        $this->created        = time();
        $this->modified       = time();
        $this->title          = '';
        $this->subject        = '';
        $this->description    = '';
        $this->keywords       = '';
        $this->category       = '';
        $this->company        = '';
        $this->manager        = '';
    }

    
    public function getCreator()
    {
        return $this->creator;
    }

    
    public function setCreator($value = '')
    {
        $this->creator = $this->setValue($value, '');

        return $this;
    }

    
    public function getLastModifiedBy()
    {
        return $this->lastModifiedBy;
    }

    
    public function setLastModifiedBy($value = '')
    {
        $this->lastModifiedBy = $this->setValue($value, $this->creator);

        return $this;
    }

    public function getCreated()
    {
        return $this->created;
    }

    public function setCreated($value = null)
    {
        $this->created = $this->setValue($value, time());

        return $this;
    }

   
    public function getModified()
    {
        return $this->modified;
    }

    
    public function setModified($value = null)
    {
        $this->modified = $this->setValue($value, time());

        return $this;
    }

    
    public function getTitle()
    {
        return $this->title;
    }

    
    public function setTitle($value = '')
    {
        $this->title = $this->setValue($value, '');

        return $this;
    }


    public function getDescription()
    {
        return $this->description;
    }

    public function setDescription($value = '')
    {
        $this->description = $this->setValue($value, '');

        return $this;
    }

    
    public function getSubject()
    {
        return $this->subject;
    }

    
    public function setSubject($value = '')
    {
        $this->subject = $this->setValue($value, '');

        return $this;
    }

    
    public function getKeywords()
    {
        return $this->keywords;
    }

    
    public function setKeywords($value = '')
    {
        $this->keywords = $this->setValue($value, '');

        return $this;
    }

   
    public function getCategory()
    {
        return $this->category;
    }

   
    public function setCategory($value = '')
    {
        $this->category = $this->setValue($value, '');

        return $this;
    }

    
    public function getCompany()
    {
        return $this->company;
    }

   
    public function setCompany($value = '')
    {
        $this->company = $this->setValue($value, '');

        return $this;
    }

    
    public function getManager()
    {
        return $this->manager;
    }

    
    public function setManager($value = '')
    {
        $this->manager = $this->setValue($value, '');

        return $this;
    }

    
    public function getCustomProperties()
    {
        return array_keys($this->customProperties);
    }

   
    public function isCustomPropertySet($propertyName)
    {
        return isset($this->customProperties[$propertyName]);
    }

   
    public function getCustomPropertyValue($propertyName)
    {
        if ($this->isCustomPropertySet($propertyName)) {
            return $this->customProperties[$propertyName]['value'];
        } else {
            return null;
        }
    }

    
    public function getCustomPropertyType($propertyName)
    {
        if ($this->isCustomPropertySet($propertyName)) {
            return $this->customProperties[$propertyName]['type'];
        } else {
            return null;
        }
    }

    
    public function setCustomProperty($propertyName, $propertyValue = '', $propertyType = null)
    {
        $propertyTypes = array(
            self::PROPERTY_TYPE_INTEGER,
            self::PROPERTY_TYPE_FLOAT,
            self::PROPERTY_TYPE_STRING,
            self::PROPERTY_TYPE_DATE,
            self::PROPERTY_TYPE_BOOLEAN
        );
        if (($propertyType === null) || (!in_array($propertyType, $propertyTypes))) {
            if ($propertyValue === null) {
                $propertyType = self::PROPERTY_TYPE_STRING;
            } elseif (is_float($propertyValue)) {
                $propertyType = self::PROPERTY_TYPE_FLOAT;
            } elseif (is_int($propertyValue)) {
                $propertyType = self::PROPERTY_TYPE_INTEGER;
            } elseif (is_bool($propertyValue)) {
                $propertyType = self::PROPERTY_TYPE_BOOLEAN;
            } else {
                $propertyType = self::PROPERTY_TYPE_STRING;
            }
        }

        $this->customProperties[$propertyName] = array(
            'value' => $propertyValue,
            'type' => $propertyType
        );
        return $this;
    }

  
    public static function convertProperty($propertyValue, $propertyType)
    {
        $conversion = self::getConversion($propertyType);

        switch ($conversion) {
            case 'empty': 
                return '';
            case 'null': 
                return null;
            case 'int': 
                return (int) $propertyValue;
            case 'uint': 
                return abs((int) $propertyValue);
            case 'float': 
                return (float) $propertyValue;
            case 'date': 
                return strtotime($propertyValue);
            case 'bool': 
                return ($propertyValue == 'true') ? true : false;
        }

        return $propertyValue;
    }

    
    public static function convertPropertyType($propertyType)
    {
        $typeGroups = array(
            self::PROPERTY_TYPE_INTEGER => array('i1', 'i2', 'i4', 'i8', 'int', 'ui1', 'ui2', 'ui4', 'ui8', 'uint'),
            self::PROPERTY_TYPE_FLOAT   => array('r4', 'r8', 'decimal'),
            self::PROPERTY_TYPE_STRING  => array('empty', 'null', 'lpstr', 'lpwstr', 'bstr'),
            self::PROPERTY_TYPE_DATE    => array('date', 'filetime'),
            self::PROPERTY_TYPE_BOOLEAN => array('bool'),
        );
        foreach ($typeGroups as $groupId => $groupMembers) {
            if (in_array($propertyType, $groupMembers)) {
                return $groupId;
            }
        }

        return self::PROPERTY_TYPE_UNKNOWN;
    }

    
    private function setValue($value, $default)
    {
        if ($value === null || $value == '') {
            $value = $default;
        }

        return $value;
    }

    
    private static function getConversion($propertyType)
    {
        $conversions = array(
            'empty' => array('empty'),
            'null'  => array('null'),
            'int'   => array('i1', 'i2', 'i4', 'i8', 'int'),
            'uint'  => array('ui1', 'ui2', 'ui4', 'ui8', 'uint'),
            'float' => array('r4', 'r8', 'decimal'),
            'bool'  => array('bool'),
            'date'  => array('date', 'filetime'),
        );
        foreach ($conversions as $conversion => $types) {
            if (in_array($propertyType, $types)) {
                return $conversion;
            }
        }

        return 'string';
    }
}
