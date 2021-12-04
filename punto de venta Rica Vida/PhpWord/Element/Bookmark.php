<?php

namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Shared\String;
use PhpOffice\PhpWord\Style;

/**
 * Bookmark element
 */
class Bookmark extends AbstractElement
{
    /**
     * Bookmark Name
     *
     * @var string
     */
    private $name;

    /**
     * Is part of collection
     *
     * @var bool
     */
    protected $collectionRelation = true;

    /**
     * Create a new Bookmark Element
     *
     * @param string $name
     */
    public function __construct($name)
    {

        $this->name = String::toUTF8($name);
        return $this;
    }

    /**
     * Get Bookmark name
     *
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }
}
