<?php


namespace PhpOffice\PhpWord\Reader\Word2007;


class DocPropsApp extends DocPropsCore
{
    
    protected $mapping = array('Company' => 'setCompany', 'Manager' => 'setManager');

    protected $callbacks = array();
}
