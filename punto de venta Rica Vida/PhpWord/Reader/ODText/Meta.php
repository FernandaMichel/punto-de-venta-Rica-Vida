<?php

namespace PhpOffice\PhpWord\Reader\ODText;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\XMLReader;


class Meta extends AbstractPart
{
    
    public function read(PhpWord $phpWord)
    {
        $xmlReader = new XMLReader();
        $xmlReader->getDomFromZip($this->docFile, $this->xmlFile);
        $docProps = $phpWord->getDocInfo();

        $metaNode = $xmlReader->getElement('office:meta');

        $properties = array(
            'title'          => 'dc:title',
            'subject'        => 'dc:subject',
            'description'    => 'dc:description',
            'keywords'       => 'meta:keyword',
            'creator'        => 'meta:initial-creator',
            'lastModifiedBy' => 'dc:creator',
        );
        foreach ($properties as $property => $path) {
            $method = "set{$property}";
            $propertyNode = $xmlReader->getElement($path, $metaNode);
            if ($propertyNode !== null && method_exists($docProps, $method)) {
                $docProps->$method($propertyNode->nodeValue);
            }
        }

        $propertyNodes = $xmlReader->getElements('meta:user-defined', $metaNode);
        foreach ($propertyNodes as $propertyNode) {
            $property = $xmlReader->getAttribute('meta:name', $propertyNode);

            if (in_array($property, array('Category', 'Company', 'Manager'))) {
                $method = "set{$property}";
                $docProps->$method($propertyNode->nodeValue);

            } else {
                $docProps->setCustomProperty($property, $propertyNode->nodeValue);
            }
        }
    }
}
