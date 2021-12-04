<?php

namespace PhpOffice\PhpWord\Style;


class Numbering extends AbstractStyle
{
    
    private $numId;

    
    private $type;

    
    private $levels = array();

    
    public function getNumId()
    {
        return $this->numId;
    }

    /**
     * Set Id
     *
     * @param integer $value
     * @return self
     */
    public function setNumId($value)
    {
        $this->numId = $this->setIntVal($value, $this->numId);

        return $this;
    }

    /**
     * Get multilevel type
     *
     * @return string
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * Set multilevel type
     *
     * @param string $value
     * @return self
     */
    public function setType($value)
    {
        $enum = array('singleLevel', 'multilevel', 'hybridMultilevel');
        $this->type = $this->setEnumVal($value, $enum, $this->type);

        return $this;
    }

    /**
     * Get levels
     *
     * @return NumberingLevel[]
     */
    public function getLevels()
    {
        return $this->levels;
    }

    /**
     * Set multilevel type
     *
     * @param array $values
     * @return self
     */
    public function setLevels($values)
    {
        if (is_array($values)) {
            foreach ($values as $key => $value) {
                $numberingLevel = new NumberingLevel();
                if (is_array($value)) {
                    $numberingLevel->setStyleByArray($value);
                    $numberingLevel->setLevel($key);
                }
                $this->levels[$key] = $numberingLevel;
            }
        }

        return $this;
    }
}
