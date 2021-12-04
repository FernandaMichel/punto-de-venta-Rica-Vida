<?php

namespace PhpOffice\PhpWord\Reader\Word2007;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\XMLReader;


abstract class AbstractPart
{
    
    const READ_VALUE = 'attributeValue';            
    const READ_EQUAL = 'attributeEquals';          
    const READ_TRUE  = 'attributeTrue';             
    const READ_FALSE = 'attributeFalse';           
    const READ_SIZE  = 'attributeMultiplyByTwo';    

    
    protected $docFile;

    
    protected $xmlFile;

    
    protected $rels = array();

    
    abstract public function read(PhpWord $phpWord);

   
    public function __construct($docFile, $xmlFile)
    {
        $this->docFile = $docFile;
        $this->xmlFile = $xmlFile;
    }

    
    public function setRels($value)
    {
        $this->rels = $value;
    }

   
    protected function readParagraph(XMLReader $xmlReader, \DOMElement $domNode, $parent, $docPart = 'document')
    {
        $paragraphStyle = null;
        $headingMatches = array();
        if ($xmlReader->elementExists('w:pPr', $domNode)) {
            $paragraphStyle = $this->readParagraphStyle($xmlReader, $domNode);
            if (is_array($paragraphStyle) && isset($paragraphStyle['styleName'])) {
                preg_match('/Heading(\d)/', $paragraphStyle['styleName'], $headingMatches);
            }
        }

        if ($xmlReader->elementExists('w:r/w:instrText', $domNode)) {
            $ignoreText = false;
            $textContent = '';
            $fontStyle = $this->readFontStyle($xmlReader, $domNode);
            $nodes = $xmlReader->getElements('w:r', $domNode);
            foreach ($nodes as $node) {
                $instrText = $xmlReader->getValue('w:instrText', $node);
                if ($xmlReader->elementExists('w:fldChar', $node)) {
                    $fldCharType = $xmlReader->getAttribute('w:fldCharType', $node, 'w:fldChar');
                    if ($fldCharType == 'begin') {
                        $ignoreText = true;
                    } elseif ($fldCharType == 'end') {
                        $ignoreText = false;
                    }
                }
                if (!is_null($instrText)) {
                    $textContent .= '{' . $instrText . '}';
                } else {
                    if ($ignoreText === false) {
                        $textContent .= $xmlReader->getValue('w:t', $node);
                    }
                }
            }
            $parent->addPreserveText($textContent, $fontStyle, $paragraphStyle);

        } elseif ($xmlReader->elementExists('w:pPr/w:numPr', $domNode)) {
            $textContent = '';
            $numId = $xmlReader->getAttribute('w:val', $domNode, 'w:pPr/w:numPr/w:numId');
            $levelId = $xmlReader->getAttribute('w:val', $domNode, 'w:pPr/w:numPr/w:ilvl');
            $nodes = $xmlReader->getElements('w:r', $domNode);
            foreach ($nodes as $node) {
                $textContent .= $xmlReader->getValue('w:t', $node);
            }
            $parent->addListItem($textContent, $levelId, null, "PHPWordList{$numId}", $paragraphStyle);

        } elseif (!empty($headingMatches)) {
            $textContent = '';
            $nodes = $xmlReader->getElements('w:r', $domNode);
            foreach ($nodes as $node) {
                $textContent .= $xmlReader->getValue('w:t', $node);
            }
            $parent->addTitle($textContent, $headingMatches[1]);

        } else {
            $runCount = $xmlReader->countElements('w:r', $domNode);
            $linkCount = $xmlReader->countElements('w:hyperlink', $domNode);
            $runLinkCount = $runCount + $linkCount;
            if ($runLinkCount == 0) {
                $parent->addTextBreak(null, $paragraphStyle);
            } else {
                $nodes = $xmlReader->getElements('*', $domNode);
                foreach ($nodes as $node) {
                    $this->readRun(
                        $xmlReader,
                        $node,
                        ($runLinkCount > 1) ? $parent->addTextRun($paragraphStyle) : $parent,
                        $docPart,
                        $paragraphStyle
                    );
                }
            }
        }
    }

   
    protected function readRun(XMLReader $xmlReader, \DOMElement $domNode, $parent, $docPart, $paragraphStyle = null)
    {
        if (!in_array($domNode->nodeName, array('w:r', 'w:hyperlink'))) {
            return;
        }
        $fontStyle = $this->readFontStyle($xmlReader, $domNode);

        if ($domNode->nodeName == 'w:hyperlink') {
            $rId = $xmlReader->getAttribute('r:id', $domNode);
            $textContent = $xmlReader->getValue('w:r/w:t', $domNode);
            $target = $this->getMediaTarget($docPart, $rId);
            if (!is_null($target)) {
                $parent->addLink($target, $textContent, $fontStyle, $paragraphStyle);
            }
        } else {
            if ($xmlReader->elementExists('w:footnoteReference', $domNode)) {
                $parent->addFootnote();

            } elseif ($xmlReader->elementExists('w:endnoteReference', $domNode)) {
                $parent->addEndnote();

            } elseif ($xmlReader->elementExists('w:pict', $domNode)) {
                $rId = $xmlReader->getAttribute('r:id', $domNode, 'w:pict/v:shape/v:imagedata');
                $target = $this->getMediaTarget($docPart, $rId);
                if (!is_null($target)) {
                    $imageSource = "zip://{$this->docFile}#{$target}";
                    $parent->addImage($imageSource);
                }

            } elseif ($xmlReader->elementExists('w:object', $domNode)) {
                $rId = $xmlReader->getAttribute('r:id', $domNode, 'w:object/o:OLEObject');
                $target = $this->getMediaTarget($docPart, $rId);
                if (!is_null($target)) {
                    $textContent = "<Object: {$target}>";
                    $parent->addText($textContent, $fontStyle, $paragraphStyle);
                }

            } else {
                $textContent = $xmlReader->getValue('w:t', $domNode);
                $parent->addText($textContent, $fontStyle, $paragraphStyle);
            }
        }
    }

    
    protected function readTable(XMLReader $xmlReader, \DOMElement $domNode, $parent, $docPart = 'document')
    {
        $tblStyle = null;
        if ($xmlReader->elementExists('w:tblPr', $domNode)) {
            $tblStyle = $this->readTableStyle($xmlReader, $domNode);
        }

        $table = $parent->addTable($tblStyle);
        $tblNodes = $xmlReader->getElements('*', $domNode);
        foreach ($tblNodes as $tblNode) {
            if ($tblNode->nodeName == 'w:tblGrid') { 

            } elseif ($tblNode->nodeName == 'w:tr') { 
                $rowHeight = $xmlReader->getAttribute('w:val', $tblNode, 'w:trPr/w:trHeight');
                $rowHRule = $xmlReader->getAttribute('w:hRule', $tblNode, 'w:trPr/w:trHeight');
                $rowHRule = $rowHRule == 'exact' ? true : false;
                $rowStyle = array(
                    'tblHeader' => $xmlReader->elementExists('w:trPr/w:tblHeader', $tblNode),
                    'cantSplit' => $xmlReader->elementExists('w:trPr/w:cantSplit', $tblNode),
                    'exactHeight' => $rowHRule,
                );

                $row = $table->addRow($rowHeight, $rowStyle);
                $rowNodes = $xmlReader->getElements('*', $tblNode);
                foreach ($rowNodes as $rowNode) {
                    if ($rowNode->nodeName == 'w:trPr') { 

                    } elseif ($rowNode->nodeName == 'w:tc') { 
                        $cellWidth = $xmlReader->getAttribute('w:w', $rowNode, 'w:tcPr/w:tcW');
                        $cellStyle = null;
                        $cellStyleNode = $xmlReader->getElement('w:tcPr', $rowNode);
                        if (!is_null($cellStyleNode)) {
                            $cellStyle = $this->readCellStyle($xmlReader, $cellStyleNode);
                        }

                        $cell = $row->addCell($cellWidth, $cellStyle);
                        $cellNodes = $xmlReader->getElements('*', $rowNode);
                        foreach ($cellNodes as $cellNode) {
                            if ($cellNode->nodeName == 'w:p') { 
                                $this->readParagraph($xmlReader, $cellNode, $cell, $docPart);
                            }
                        }
                    }
                }
            }
        }
    }

    
    protected function readParagraphStyle(XMLReader $xmlReader, \DOMElement $domNode)
    {
        if (!$xmlReader->elementExists('w:pPr', $domNode)) {
            return null;
        }

        $styleNode = $xmlReader->getElement('w:pPr', $domNode);
        $styleDefs = array(
            'styleName'       => array(self::READ_VALUE, 'w:pStyle'),
            'align'           => array(self::READ_VALUE, 'w:jc'),
            'basedOn'         => array(self::READ_VALUE, 'w:basedOn'),
            'next'            => array(self::READ_VALUE, 'w:next'),
            'indent'          => array(self::READ_VALUE, 'w:ind', 'w:left'),
            'hanging'         => array(self::READ_VALUE, 'w:ind', 'w:hanging'),
            'spaceAfter'      => array(self::READ_VALUE, 'w:spacing', 'w:after'),
            'spaceBefore'     => array(self::READ_VALUE, 'w:spacing', 'w:before'),
            'widowControl'    => array(self::READ_FALSE, 'w:widowControl'),
            'keepNext'        => array(self::READ_TRUE,  'w:keepNext'),
            'keepLines'       => array(self::READ_TRUE,  'w:keepLines'),
            'pageBreakBefore' => array(self::READ_TRUE,  'w:pageBreakBefore'),
        );

        return $this->readStyleDefs($xmlReader, $styleNode, $styleDefs);
    }

    
    protected function readFontStyle(XMLReader $xmlReader, \DOMElement $domNode)
    {
        if (is_null($domNode)) {
            return null;
        }
        if ($domNode->nodeName == 'w:hyperlink') {
            $domNode = $xmlReader->getElement('w:r', $domNode);
        }
        if (!$xmlReader->elementExists('w:rPr', $domNode)) {
            return null;
        }

        $styleNode = $xmlReader->getElement('w:rPr', $domNode);
        $styleDefs = array(
            'styleName'           => array(self::READ_VALUE, 'w:rStyle'),
            'name'                => array(self::READ_VALUE, 'w:rFonts', 'w:ascii'),
            'hint'                => array(self::READ_VALUE, 'w:rFonts', 'w:hint'),
            'size'                => array(self::READ_SIZE,  'w:sz'),
            'color'               => array(self::READ_VALUE, 'w:color'),
            'underline'           => array(self::READ_VALUE, 'w:u'),
            'bold'                => array(self::READ_TRUE,  'w:b'),
            'italic'              => array(self::READ_TRUE,  'w:i'),
            'strikethrough'       => array(self::READ_TRUE,  'w:strike'),
            'doubleStrikethrough' => array(self::READ_TRUE,  'w:dstrike'),
            'smallCaps'           => array(self::READ_TRUE,  'w:smallCaps'),
            'allCaps'             => array(self::READ_TRUE,  'w:caps'),
            'superScript'         => array(self::READ_EQUAL, 'w:vertAlign', 'w:val', 'superscript'),
            'subScript'           => array(self::READ_EQUAL, 'w:vertAlign', 'w:val', 'subscript'),
            'fgColor'             => array(self::READ_VALUE, 'w:highlight'),
            'rtl'                 => array(self::READ_TRUE,  'w:rtl'),
        );

        return $this->readStyleDefs($xmlReader, $styleNode, $styleDefs);
    }

    
    protected function readTableStyle(XMLReader $xmlReader, \DOMElement $domNode)
    {
        $style = null;
        $margins = array('top', 'left', 'bottom', 'right');
        $borders = $margins + array('insideH', 'insideV');

        if ($xmlReader->elementExists('w:tblPr', $domNode)) {
            if ($xmlReader->elementExists('w:tblPr/w:tblStyle', $domNode)) {
                $style = $xmlReader->getAttribute('w:val', $domNode, 'w:tblPr/w:tblStyle');
            } else {
                $styleNode = $xmlReader->getElement('w:tblPr', $domNode);
                $styleDefs = array();
                foreach ($margins as $side) {
                    $ucfSide = ucfirst($side);
                    $styleDefs["cellMargin$ucfSide"] = array(self::READ_VALUE, "w:tblCellMar/w:$side", 'w:w');
                }
                foreach ($borders as $side) {
                    $ucfSide = ucfirst($side);
                    $styleDefs["border{$ucfSide}Size"] = array(self::READ_VALUE, "w:tblBorders/w:$side", 'w:sz');
                    $styleDefs["border{$ucfSide}Color"] = array(self::READ_VALUE, "w:tblBorders/w:$side", 'w:color');
                }
                $style = $this->readStyleDefs($xmlReader, $styleNode, $styleDefs);
            }
        }

        return $style;
    }

    
    private function readCellStyle(XMLReader $xmlReader, \DOMElement $domNode)
    {
        $styleDefs = array(
            'valign'        => array(self::READ_VALUE, 'w:vAlign'),
            'textDirection' => array(self::READ_VALUE, 'w:textDirection'),
            'gridSpan'      => array(self::READ_VALUE, 'w:gridSpan'),
            'vMerge'        => array(self::READ_VALUE, 'w:vMerge'),
            'bgColor'       => array(self::READ_VALUE, 'w:shd/w:fill'),
        );

        return $this->readStyleDefs($xmlReader, $domNode, $styleDefs);
    }

    
    protected function readStyleDefs(XMLReader $xmlReader, \DOMElement $parentNode = null, $styleDefs = array())
    {
        $styles = array();

        foreach ($styleDefs as $styleProp => $styleVal) {
            @list($method, $element, $attribute, $expected) = $styleVal;

            if ($xmlReader->elementExists($element, $parentNode)) {
                $node = $xmlReader->getElement($element, $parentNode);

                $attribute = ($attribute === null) ? 'w:val' : $attribute;
                $attributeValue = $xmlReader->getAttribute($attribute, $node);

                $styleValue = $this->readStyleDef($method, $attributeValue, $expected);
                if ($styleValue !== null) {
                    $styles[$styleProp] = $styleValue;
                }
            }
        }

        return $styles;
    }

    
    private function readStyleDef($method, $attributeValue, $expected)
    {
        $style = $attributeValue;

        if ($method == self::READ_SIZE) {
            $style = $attributeValue / 2;
        } elseif ($method == self::READ_TRUE) {
            $style = true;
        } elseif ($method == self::READ_FALSE) {
            $style = false;
        } elseif ($method == self::READ_EQUAL) {
            $style = $attributeValue == $expected;
        }

        return $style;
    }

    
    private function getMediaTarget($docPart, $rId)
    {
        $target = null;

        if (isset($this->rels[$docPart]) && isset($this->rels[$docPart][$rId])) {
            $target = $this->rels[$docPart][$rId]['target'];
        }

        return $target;
    }
}
