<?php
namespace PhpOffice\PhpWord\Reader\Word2007;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\XMLReader;


class Styles extends AbstractPart
{
   
    public function read(PhpWord $phpWord)
    {
        $xmlReader = new XMLReader();
        $xmlReader->getDomFromZip($this->docFile, $this->xmlFile);

        $nodes = $xmlReader->getElements('w:style');
        if ($nodes->length > 0) {
            foreach ($nodes as $node) {
                $type = $xmlReader->getAttribute('w:type', $node);
                $name = $xmlReader->getAttribute('w:styleId', $node);
                if (is_null($name)) {
                    $name = $xmlReader->getAttribute('w:val', $node, 'w:name');
                }
                preg_match('/Heading(\d)/', $name, $headingMatches);
                switch ($type) {

                    case 'paragraph':
                        $paragraphStyle = $this->readParagraphStyle($xmlReader, $node);
                        $fontStyle = $this->readFontStyle($xmlReader, $node);
                        if (!empty($headingMatches)) {
                            $phpWord->addTitleStyle($headingMatches[1], $fontStyle, $paragraphStyle);
                        } else {
                            if (empty($fontStyle)) {
                                if (is_array($paragraphStyle)) {
                                    $phpWord->addParagraphStyle($name, $paragraphStyle);
                                }
                            } else {
                                $phpWord->addFontStyle($name, $fontStyle, $paragraphStyle);
                            }
                        }
                        break;

                    case 'character':
                        $fontStyle = $this->readFontStyle($xmlReader, $node);
                        if (!empty($fontStyle)) {
                            $phpWord->addFontStyle($name, $fontStyle);
                        }
                        break;

                    case 'table':
                        $tStyle = $this->readTableStyle($xmlReader, $node);
                        if (!empty($tStyle)) {
                            $phpWord->addTableStyle($name, $tStyle);
                        }
                        break;
                }
            }
        }
    }
}
