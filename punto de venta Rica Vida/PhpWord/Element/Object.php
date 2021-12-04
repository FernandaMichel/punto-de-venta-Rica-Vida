<?php

namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Exception\InvalidObjectException;
use PhpOffice\PhpWord\Style\Image as ImageStyle;


class Object extends AbstractElement
{
    
    private $source;

    
    private $style;

    
    private $icon;

    
    private $imageRelationId;

    
    protected $mediaRelation = true;

    
    public function __construct($source, $style = null)
    {
        $supportedTypes = array('xls', 'doc', 'ppt', 'xlsx', 'docx', 'pptx');
        $pathInfo = pathinfo($source);

        if (file_exists($source) && in_array($pathInfo['extension'], $supportedTypes)) {
            $ext = $pathInfo['extension'];
            if (strlen($ext) == 4 && strtolower(substr($ext, -1)) == 'x') {
                $ext = substr($ext, 0, -1);
            }

            $this->source = $source;
            $this->style = $this->setNewStyle(new ImageStyle(), $style, true);
            $this->icon = realpath(__DIR__ . "/../resources/{$ext}.png");

            return $this;
        } else {
            throw new InvalidObjectException();
        }
    }

    
    public function getSource()
    {
        return $this->source;
    }


    public function getStyle()
    {
        return $this->style;
    }

    public function getIcon()
    {
        return $this->icon;
    }

    
    public function getImageRelationId()
    {
        return $this->imageRelationId;
    }

    
    public function setImageRelationId($rId)
    {
        $this->imageRelationId = $rId;
    }

    
    public function getObjectId()
    {
        return $this->relationId + 1325353440;
    }

  
    public function setObjectId($objId)
    {
        $this->relationId = $objId;
    }
}
