<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Style\Paragraph;

class TextRun extends AbstractContainer
{
   
    protected $container = 'TextRun';

    
    protected $paragraphStyle;

    
    public function __construct($paragraphStyle = null)
    {
        $this->paragraphStyle = $this->setNewStyle(new Paragraph(), $paragraphStyle);
    }

    
    public function getParagraphStyle()
    {
        return $this->paragraphStyle;
    }
}
