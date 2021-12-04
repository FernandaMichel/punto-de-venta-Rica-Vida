<?php


namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\Exception\InvalidImageException;
use PhpOffice\PhpWord\Exception\UnsupportedImageTypeException;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\Shared\ZipArchive;
use PhpOffice\PhpWord\Style\Image as ImageStyle;

class Image extends AbstractElement
{
   
    const SOURCE_LOCAL = 'local'; 
    const SOURCE_GD = 'gd'; 
    const SOURCE_ARCHIVE = 'archive'; 
    
    private $source;

   
    private $sourceType;

   
    private $style;

    
    private $watermark;

    
    private $imageType;

    private $imageCreateFunc;

    
    private $imageFunc;

    
    private $imageExtension;

    
    private $memoryImage;

    
    private $target;

    
    private $mediaIndex;

    
    protected $mediaRelation = true;

    
    public function __construct($source, $style = null, $watermark = false)
    {
        $this->source = $source;
        $this->setIsWatermark($watermark);
        $this->style = $this->setNewStyle(new ImageStyle(), $style, true);

        $this->checkImage($source);
    }

    
    public function getStyle()
    {
        return $this->style;
    }

    
    public function getSource()
    {
        return $this->source;
    }

    
    public function getSourceType()
    {
        return $this->sourceType;
    }

   
    public function getMediaId()
    {
        return md5($this->source);
    }

    
    public function isWatermark()
    {
        return $this->watermark;
    }

    public function setIsWatermark($value)
    {
        $this->watermark = $value;
    }

    
    public function getImageType()
    {
        return $this->imageType;
    }

   
    public function getImageCreateFunction()
    {
        return $this->imageCreateFunc;
    }

    
    public function getImageFunction()
    {
        return $this->imageFunc;
    }

    
    public function getImageExtension()
    {
        return $this->imageExtension;
    }

    
    public function isMemImage()
    {
        return $this->memoryImage;
    }

    public function getTarget()
    {
        return $this->target;
    }

    
    public function setTarget($value)
    {
        $this->target = $value;
    }

    
    public function getMediaIndex()
    {
        return $this->mediaIndex;
    }

    
    public function setMediaIndex($value)
    {
        $this->mediaIndex = $value;
    }

    
    public function getImageStringData($base64 = false)
    {
        $source = $this->source;
        $actualSource = null;
        $imageBinary = null;
        $imageData = null;
        $isTemp = false;

        
        if ($this->sourceType == self::SOURCE_ARCHIVE) {
            $source = substr($source, 6);
            list($zipFilename, $imageFilename) = explode('#', $source);

            $zip = new ZipArchive();
            if ($zip->open($zipFilename) !== false) {
                if ($zip->locateName($imageFilename)) {
                    $isTemp = true;
                    $zip->extractTo(Settings::getTempDir(), $imageFilename);
                    $actualSource = Settings::getTempDir() . DIRECTORY_SEPARATOR . $imageFilename;
                }
            }
            $zip->close();
        } else {
            $actualSource = $source;
        }

       
        if ($this->sourceType == self::SOURCE_GD) {
            $imageResource = call_user_func($this->imageCreateFunc, $actualSource);
            ob_start();
            call_user_func($this->imageFunc, $imageResource);
            $imageBinary = ob_get_contents();
            ob_end_clean();
        } else {
            $fileHandle = fopen($actualSource, 'rb', false);
            if ($fileHandle !== false) {
                $imageBinary = fread($fileHandle, filesize($actualSource));
                fclose($fileHandle);
            }
        }
        if ($imageBinary !== null) {
            if ($base64) {
                $imageData = chunk_split(base64_encode($imageBinary));
            } else {
                $imageData = chunk_split(bin2hex($imageBinary));
            }
        }

        if ($isTemp === true) {
            @unlink($actualSource);
        }

        return $imageData;
    }

    
    private function checkImage($source)
    {
        $this->setSourceType($source);

        if ($this->sourceType == self::SOURCE_ARCHIVE) {
            $imageData = $this->getArchiveImageSize($source);
        } else {
            $imageData = @getimagesize($source);
        }
        if (!is_array($imageData)) {
            throw new InvalidImageException();
        }
        list($actualWidth, $actualHeight, $imageType) = $imageData;

        $supportedTypes = array(IMAGETYPE_JPEG, IMAGETYPE_GIF, IMAGETYPE_PNG);
        if ($this->sourceType != self::SOURCE_GD) {
            $supportedTypes = array_merge($supportedTypes, array(IMAGETYPE_BMP, IMAGETYPE_TIFF_II, IMAGETYPE_TIFF_MM));
        }
        if (!in_array($imageType, $supportedTypes)) {
            throw new UnsupportedImageTypeException();
        }

        $this->imageType = image_type_to_mime_type($imageType);
        $this->setFunctions();
        $this->setProportionalSize($actualWidth, $actualHeight);
    }

    
    private function setSourceType($source)
    {
        if (stripos(strrev($source), strrev('.php')) === 0) {
            $this->memoryImage = true;
            $this->sourceType = self::SOURCE_GD;
        } elseif (strpos($source, 'zip://') !== false) {
            $this->memoryImage = false;
            $this->sourceType = self::SOURCE_ARCHIVE;
        } else {
            $this->memoryImage = (filter_var($source, FILTER_VALIDATE_URL) !== false);
            $this->sourceType = $this->memoryImage ? self::SOURCE_GD : self::SOURCE_LOCAL;
        }
    }

    
    private function getArchiveImageSize($source)
    {
        $imageData = null;
        $source = substr($source, 6);
        list($zipFilename, $imageFilename) = explode('#', $source);

        $tempFilename = tempnam(Settings::getTempDir(), 'PHPWordImage');
        if (false === $tempFilename) {
            throw new CreateTemporaryFileException();
        }

        $zip = new ZipArchive();
        if ($zip->open($zipFilename) !== false) {
            if ($zip->locateName($imageFilename)) {
                $imageContent = $zip->getFromName($imageFilename);
                if ($imageContent !== false) {
                    file_put_contents($tempFilename, $imageContent);
                    $imageData = getimagesize($tempFilename);
                    unlink($tempFilename);
                }
            }
            $zip->close();
        }

        return $imageData;
    }

    
    private function setFunctions()
    {
        switch ($this->imageType) {
            case 'image/png':
                $this->imageCreateFunc = 'imagecreatefrompng';
                $this->imageFunc = 'imagepng';
                $this->imageExtension = 'png';
                break;
            case 'image/gif':
                $this->imageCreateFunc = 'imagecreatefromgif';
                $this->imageFunc = 'imagegif';
                $this->imageExtension = 'gif';
                break;
            case 'image/jpeg':
            case 'image/jpg':
                $this->imageCreateFunc = 'imagecreatefromjpeg';
                $this->imageFunc = 'imagejpeg';
                $this->imageExtension = 'jpg';
                break;
            case 'image/bmp':
            case 'image/x-ms-bmp':
                $this->imageType = 'image/bmp';
                $this->imageExtension = 'bmp';
                break;
            case 'image/tiff':
                $this->imageExtension = 'tif';
                break;
        }
    }

    
    private function setProportionalSize($actualWidth, $actualHeight)
    {
        $styleWidth = $this->style->getWidth();
        $styleHeight = $this->style->getHeight();
        if (!($styleWidth && $styleHeight)) {
            if ($styleWidth == null && $styleHeight == null) {
                $this->style->setWidth($actualWidth);
                $this->style->setHeight($actualHeight);
            } elseif ($styleWidth) {
                $this->style->setHeight($actualHeight * ($styleWidth / $actualWidth));
            } else {
                $this->style->setWidth($actualWidth * ($styleHeight / $actualHeight));
            }
        }
    }

    
    public function getIsWatermark()
    {
        return $this->isWatermark();
    }

    
    public function getIsMemImage()
    {
        return $this->isMemImage();
    }
}
