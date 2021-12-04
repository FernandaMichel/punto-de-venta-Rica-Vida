<?php


namespace PhpOffice\PhpWord\Element;

/**
 * Footer element
 */
class Footer extends AbstractContainer
{
    /**
     * Header/footer types constants
     *
     * @var string
     * @link http://www.schemacentral.com/sc/ooxml/a-wtype-4.html Header or Footer Type
     */
    const AUTO  = 'default';  // default and odd pages
    const FIRST = 'first';
    const EVEN  = 'even';

    /**
     * @var string Container type
     */
    protected $container = 'Footer';

    /**
     * Header type
     *
     * @var string
     */
    protected $type = self::AUTO;

    /**
     * Create new instance
     *
     * @param int $sectionId
     * @param int $containerId
     * @param string $type
     */
    public function __construct($sectionId, $containerId = 1, $type = self::AUTO)
    {
        $this->sectionId = $sectionId;
        $this->setType($type);
        $this->setDocPart($this->container, ($sectionId - 1) * 3 + $containerId);
    }

    /**
     * Set type.
     *
     * @since 0.10.0
     *
     * @param string $value
     * @return void
     */
    public function setType($value = self::AUTO)
    {
        if (!in_array($value, array(self::AUTO, self::FIRST, self::EVEN))) {
            $value = self::AUTO;
        }
        $this->type = $value;
    }

    /**
     * Get type
     *
     * @return string
     * @since 0.10.0
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * Reset type to default
     *
     * @return string
     */
    public function resetType()
    {
        return $this->type = self::AUTO;
    }

    /**
     * First page only header
     *
     * @return string
     */
    public function firstPage()
    {
        return $this->type = self::FIRST;
    }

    /**
     * Even numbered pages only
     *
     * @return string
     */
    public function evenPage()
    {
        return $this->type = self::EVEN;
    }
}
