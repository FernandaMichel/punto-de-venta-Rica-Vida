<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Shared\String;

/**
 * Check box element
 *
 * @since 0.10.0
 */
class CheckBox extends Text
{
    /**
     * Name content
     *
     * @var string
     */
    private $name;

    /**
     * Create new instance
     *
     * @param string $name
     * @param string $text
     * @param mixed $fontStyle
     * @param mixed $paragraphStyle
     * @return self
     */
    public function __construct($name = null, $text = null, $fontStyle = null, $paragraphStyle = null)
    {
        $this->setName($name);
        parent::__construct($text, $fontStyle, $paragraphStyle);
    }

    /**
     * Set name content
     *
     * @param string $name
     * @return self
     */
    public function setName($name)
    {
        $this->name = String::toUTF8($name);

        return $this;
    }

    /**
     * Get name content
     *
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }
}
