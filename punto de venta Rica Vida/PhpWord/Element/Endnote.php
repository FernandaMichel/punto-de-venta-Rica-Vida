<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Style\Paragraph;

/**
 * Endnote element
 *
 * @since 0.10.0
 */
class Endnote extends Footnote
{
    /**
     * @var string Container type
     */
    protected $container = 'Endnote';

    /**
     * Create new instance
     *
     * @param string|array|\PhpOffice\PhpWord\Style\Paragraph $paragraphStyle
     */
    public function __construct($paragraphStyle = null)
    {
        $this->paragraphStyle = $this->setNewStyle(new Paragraph(), $paragraphStyle);
    }
}
