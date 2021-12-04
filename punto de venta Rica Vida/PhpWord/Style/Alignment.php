<?php


namespace PhpOffice\PhpWord\Style;


class Alignment extends AbstractStyle
{
    
    const ALIGN_LEFT = 'left'; // Align left
    const ALIGN_RIGHT = 'right'; // Align right
    const ALIGN_CENTER = 'center'; // Align center
    const ALIGN_BOTH = 'both'; // Align both
    const ALIGN_JUSTIFY = 'justify'; // Alias for align both

    
    private $value = null;

    public function __construct($style = array())
    {
        $this->setStyleByArray($style);
    }

    
    public function getValue()
    {
        return $this->value;
    }

   
    public function setValue($value = null)
    {
        if (strtolower($value) == self::ALIGN_JUSTIFY) {
            $value = self::ALIGN_BOTH;
        }
        $enum = array(self::ALIGN_LEFT, self::ALIGN_RIGHT, self::ALIGN_CENTER, self::ALIGN_BOTH, self::ALIGN_JUSTIFY);
        $this->value = $this->setEnumVal($value, $enum, $this->value);

        return $this;
    }
}
