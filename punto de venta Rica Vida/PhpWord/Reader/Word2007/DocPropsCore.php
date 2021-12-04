<?php


namespace PhpOffice\PhpWord\Reader\Word2007;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\XMLReader;


class DocPropsCore extends AbstractPart
{
    
    protected $mapping = array(
        'dc:creator' => 'setCreator',
        'dc:title' => 'setTitle',
        'dc:description' => 'setDescription',
        'dc:subject' => 'setSubject',
        'cp:keywords' => 'setKeywords',
        'cp:category' => 'setCategory',
        'cp:lastModifiedBy' => 'setLastModifiedBy',
        'dcterms:created' => 'setCreated',
        'dcterms:modified' => 'setModified',
    );

    
    protected $callbacks = array('dcterms:created' => 'strtotime', 'dcterms:modified' => 'strtotime');

    
    public function read(PhpWord $phpWord)
    {
        $xmlReader = new XMLReader();
        $xmlReader->getDomFromZip($this->docFile, $this->xmlFile);

        $docProps = $phpWord->getDocInfo();

        $nodes = $xmlReader->getElements('*');
        if ($nodes->length > 0) {
            foreach ($nodes as $node) {
                if (!isset($this->mapping[$node->nodeName])) {
                    continue;
                }
                $method = $this->mapping[$node->nodeName];
                $value = $node->nodeValue == '' ? null : $node->nodeValue;
                if (isset($this->callbacks[$node->nodeName])) {
                    $value = $this->callbacks[$node->nodeName]($value);
                }
                if (method_exists($docProps, $method)) {
                    $docProps->$method($value);
                }
            }
        }
    }
}
