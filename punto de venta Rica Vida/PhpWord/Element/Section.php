<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\Style\Section as SectionStyle;


class Section extends AbstractContainer
{
    
    protected $container = 'Section';

    
    private $style;

    
    private $headers = array();

    
    private $footers = array();

    
    public function __construct($sectionCount, $style = null)
    {
        $this->sectionId = $sectionCount;
        $this->setDocPart($this->container, $this->sectionId);
        $this->style = new SectionStyle();
        $this->setStyle($style);
    }

    
    public function setStyle($style = null)
    {
        if (!is_null($style) && is_array($style)) {
            $this->style->setStyleByArray($style);
        }
    }

    
    public function getStyle()
    {
        return $this->style;
    }

    
    public function addHeader($type = Header::AUTO)
    {
        return $this->addHeaderFooter($type, true);
    }

    
    public function addFooter($type = Header::AUTO)
    {
        return $this->addHeaderFooter($type, false);
    }

    
    public function getHeaders()
    {
        return $this->headers;
    }

   
    public function getFooters()
    {
        return $this->footers;
    }

    
    public function hasDifferentFirstPage()
    {
        foreach ($this->headers as $header) {
            if ($header->getType() == Header::FIRST) {
                return true;
            }
        }
        return false;
    }

    
    private function addHeaderFooter($type = Header::AUTO, $header = true)
    {
        $containerClass = substr(get_class($this), 0, strrpos(get_class($this), '\\')) . '\\' .
            ($header ? 'Header' : 'Footer');
        $collectionArray = $header ? 'headers' : 'footers';
        $collection = &$this->$collectionArray;

        if (in_array($type, array(Header::AUTO, Header::FIRST, Header::EVEN))) {
            $index = count($collection);
            $container = new $containerClass($this->sectionId, ++$index, $type);
            $container->setPhpWord($this->phpWord);

            $collection[$index] = $container;
            return $container;
        } else {
            throw new Exception('Invalid header/footer type.');
        }

    }

    
    public function setSettings($settings = null)
    {
        $this->setStyle($settings);
    }

    public function getSettings()
    {
        return $this->getStyle();
    }

    
    public function createHeader()
    {
        return $this->addHeader();
    }

    
    public function createFooter()
    {
        return $this->addFooter();
    }

    
    public function getFooter()
    {
        if (empty($this->footers)) {
            return null;
        } else {
            return $this->footers[1];
        }
    }
}
