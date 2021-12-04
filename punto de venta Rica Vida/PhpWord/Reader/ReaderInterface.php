<?php


namespace PhpOffice\PhpWord\Reader;
interface ReaderInterface
{
   
    public function canRead($filename);

    
    public function load($filename);
}
