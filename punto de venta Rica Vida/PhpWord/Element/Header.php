<?php


namespace PhpOffice\PhpWord\Element;


class Header extends Footer
{
    protected $container = 'Header';

    
    public function addWatermark($src, $style = null)
    {
        return $this->addImage($src, $style, true);
    }
}
