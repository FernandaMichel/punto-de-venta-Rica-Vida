<?php


namespace PhpOffice\PhpWord\Reader\RTF;

use PhpOffice\PhpWord\PhpWord;


class Document
{
    const PARA = 'readParagraph';
    const STYL = 'readStyle';
    const SKIP = 'readSkip';

    
    private $phpWord;

    
    private $section;

    
    private $textrun;

    
    public $rtf;

    
    private $length = 0;

    
    private $offset = 0;

    private $control = '';

    
    private $text = '';

    
    private $isControl = false;

    
    private $isFirst = false;

    
    private $groups = array();

    
    private $flags = array();

    
    public function read(PhpWord $phpWord)
    {
        $markers = array(
            123 => 'markOpening',   
            125 => 'markClosing',   
            92  => 'markBackslash', 
            10  => 'markNewline',   
            13  => 'markNewline'    
        );

        $this->phpWord = $phpWord;
        $this->section = $phpWord->addSection();
        $this->textrun = $this->section->addTextRun();
        $this->length = strlen($this->rtf);

        $this->flags['paragraph'] = true; 
        while ($this->offset < $this->length) {
            $char  = $this->rtf[$this->offset];
            $ascii = ord($char);

            if (isset($markers[$ascii])) { 
                $markerFunction = $markers[$ascii];
                $this->$markerFunction();
            } else {
                if ($this->isControl === false) { 
                    $this->pushText($char);
                } else {
                    if (preg_match("/^[a-zA-Z0-9-]?$/", $char)) { 
                        $this->control .= $char;
                        $this->isFirst = false;
                    } else { 
                        if ($this->isFirst) {
                            $this->isFirst = false;
                        } else {
                            if ($char == ' ') { 
                                $this->flushControl(true);
                            }
                        }
                    }
                }
            }
            $this->offset++;
        }
        $this->flushText();
    }

   
    private function markOpening()
    {
        $this->flush(true);
        array_push($this->groups, $this->flags);
    }

    
    private function markClosing()
    {
        $this->flush(true);
        $this->flags = array_pop($this->groups);
    }

    
    private function markBackslash()
    {
        if ($this->isFirst) {
            $this->setControl(false);
            $this->text .= '\\';
        } else {
            $this->flush();
            $this->setControl(true);
            $this->control = '';
        }
    }

    
    private function markNewline()
    {
        if ($this->isControl) {
            $this->flushControl(true);
        }
    }

    
    private function flush($isControl = false)
    {
        if ($this->isControl) {
            $this->flushControl($isControl);
        } else {
            $this->flushText();
        }
    }

    
    private function flushControl($isControl = false)
    {
        if (preg_match("/^([A-Za-z]+)(-?[0-9]*) ?$/", $this->control, $match) === 1) {
            list(, $control, $parameter) = $match;
            $this->parseControl($control, $parameter);
        }

        if ($isControl === true) {
            $this->setControl(false);
        }
    }

   
    private function flushText()
    {
        if ($this->text != '') {
            if (isset($this->flags['property'])) { 
                $this->flags['value'] = $this->text;
            } else { 
                if ($this->flags['paragraph'] === true) {
                    $this->flags['paragraph'] = false;
                    $this->flags['text'] = $this->text;
                }
            }

            if (!isset($this->flags['skipped'])) {
                $this->readText();
            }

            $this->text = '';
        }
    }

   
    private function setControl($value)
    {
        $this->isControl = $value;
        $this->isFirst = $value;
    }

    
    private function pushText($char)
    {
        if ($char == '<') {
            $this->text .= "&lt;";
        } elseif ($char == '>') {
            $this->text .= "&gt;";
        } else {
            $this->text .= $char;
        }
    }

   
    private function parseControl($control, $parameter)
    {
        $controls = array(
            'par'       => array(self::PARA,    'paragraph',    true),
            'b'         => array(self::STYL,    'font',         'bold',         true),
            'i'         => array(self::STYL,    'font',         'italic',       true),
            'u'         => array(self::STYL,    'font',         'underline',    true),
            'strike'    => array(self::STYL,    'font',         'strikethrough',true),
            'fs'        => array(self::STYL,    'font',         'size',         $parameter),
            'qc'        => array(self::STYL,    'paragraph',    'align',        'center'),
            'sa'        => array(self::STYL,    'paragraph',    'spaceAfter',   $parameter),
            'fonttbl'   => array(self::SKIP,    'fonttbl',      null),
            'colortbl'  => array(self::SKIP,    'colortbl',     null),
            'info'      => array(self::SKIP,    'info',         null),
            'generator' => array(self::SKIP,    'generator',    null),
            'title'     => array(self::SKIP,    'title',        null),
            'subject'   => array(self::SKIP,    'subject',      null),
            'category'  => array(self::SKIP,    'category',     null),
            'keywords'  => array(self::SKIP,    'keywords',     null),
            'comment'   => array(self::SKIP,    'comment',      null),
            'shppict'   => array(self::SKIP,    'pic',          null),
            'fldinst'   => array(self::SKIP,    'link',         null),
        );

        if (isset($controls[$control])) {
            list($function) = $controls[$control];
            if (method_exists($this, $function)) {
                $directives = $controls[$control];
                array_shift($directives); 
                $this->$function($directives);
            }
        }
    }

    
    private function readParagraph($directives)
    {
        list($property, $value) = $directives;
        $this->textrun = $this->section->addTextRun();
        $this->flags[$property] = $value;
    }

    
    private function readStyle($directives)
    {
        list($style, $property, $value) = $directives;
        $this->flags['styles'][$style][$property] = $value;
    }

   
    private function readSkip($directives)
    {
        list($property) = $directives;
        $this->flags['property'] = $property;
        $this->flags['skipped'] = true;
    }

  
    private function readText()
    {
        $text = $this->textrun->addText($this->text);
        if (isset($this->flags['styles']['font'])) {
            $text->getFontStyle()->setStyleByArray($this->flags['styles']['font']);
        }
    }
}
