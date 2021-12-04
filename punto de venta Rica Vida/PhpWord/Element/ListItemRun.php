<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Style\ListItem as ListItemStyle;
use PhpOffice\PhpWord\Style\Paragraph;


class ListItemRun extends TextRun
{
    
    protected $container = 'ListItemRun';

    
    private $style;

   
    private $depth;

   
    public function __construct($depth = 0, $listStyle = null, $paragraphStyle = null)
    {
        $this->depth = $depth;

        if (!is_null($listStyle) && is_string($listStyle)) {
            $this->style = new ListItemStyle($listStyle);
        } else {
            $this->style = $this->setNewStyle(new ListItemStyle(), $listStyle, true);
        }
        $this->paragraphStyle = $this->setNewStyle(new Paragraph(), $paragraphStyle);
    }

  
    public function getStyle()
    {
        return $this->style;
    }

     
    public function getDepth()
    {
        return $this->depth;
    }
}
