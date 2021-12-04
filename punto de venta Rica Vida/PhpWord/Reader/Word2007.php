<?php


namespace PhpOffice\PhpWord\Reader;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\XMLReader;
use PhpOffice\PhpWord\Shared\ZipArchive;


class Word2007 extends AbstractReader implements ReaderInterface
{
    
    public function load($docFile)
    {
        $phpWord = new PhpWord();
        $relationships = $this->readRelationships($docFile);

        $steps = array(
            array('stepPart' => 'document', 'stepItems' => array(
                'styles'    => 'Styles',
                'numbering' => 'Numbering',
            )),
            array('stepPart' => 'main', 'stepItems' => array(
                'officeDocument'      => 'Document',
                'core-properties'     => 'DocPropsCore',
                'extended-properties' => 'DocPropsApp',
                'custom-properties'   => 'DocPropsCustom',
            )),
            array('stepPart' => 'document', 'stepItems' => array(
                'endnotes'  => 'Endnotes',
                'footnotes' => 'Footnotes',
            )),
        );

        foreach ($steps as $step) {
            $stepPart = $step['stepPart'];
            $stepItems = $step['stepItems'];
            foreach ($relationships[$stepPart] as $relItem) {
                $relType = $relItem['type'];
                if (isset($stepItems[$relType])) {
                    $partName = $stepItems[$relType];
                    $xmlFile = $relItem['target'];
                    $this->readPart($phpWord, $relationships, $partName, $docFile, $xmlFile);
                }
            }
        }

        return $phpWord;
    }

    
    private function readPart(PhpWord $phpWord, $relationships, $partName, $docFile, $xmlFile)
    {
        $partClass = "PhpOffice\\PhpWord\\Reader\\Word2007\\{$partName}";
        if (class_exists($partClass)) {
            $part = new $partClass($docFile, $xmlFile);
            $part->setRels($relationships);
            $part->read($phpWord);
        }

    }

    
    private function readRelationships($docFile)
    {
        $relationships = array();

        $relationships['main'] = $this->getRels($docFile, '_rels/.rels');

        $wordRelsPath = 'word/_rels/';
        $zip = new ZipArchive();
        if ($zip->open($docFile) === true) {
            for ($i = 0; $i < $zip->numFiles; $i++) {
                $xmlFile = $zip->getNameIndex($i);
                if ((substr($xmlFile, 0, strlen($wordRelsPath))) == $wordRelsPath && (substr($xmlFile, -1)) != '/') {
                    $docPart = str_replace('.xml.rels', '', str_replace($wordRelsPath, '', $xmlFile));
                    $relationships[$docPart] = $this->getRels($docFile, $xmlFile, 'word/');
                }
            }
            $zip->close();
        }

        return $relationships;
    }

    
    private function getRels($docFile, $xmlFile, $targetPrefix = '')
    {
        $metaPrefix = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/';
        $officePrefix = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

        $rels = array();

        $xmlReader = new XMLReader();
        $xmlReader->getDomFromZip($docFile, $xmlFile);
        $nodes = $xmlReader->getElements('*');
        foreach ($nodes as $node) {
            $rId = $xmlReader->getAttribute('Id', $node);
            $type = $xmlReader->getAttribute('Type', $node);
            $target = $xmlReader->getAttribute('Target', $node);

            $type = str_replace($metaPrefix, '', $type);
            $type = str_replace($officePrefix, '', $type);
            $docPart = str_replace('.xml', '', $target);

            if (!in_array($type, array('hyperlink'))) {
                $target = $targetPrefix . $target;
            }

            $rels[$rId] = array('type' => $type, 'target' => $target, 'docPart' => $docPart);
        }
        ksort($rels);

        return $rels;
    }
}
