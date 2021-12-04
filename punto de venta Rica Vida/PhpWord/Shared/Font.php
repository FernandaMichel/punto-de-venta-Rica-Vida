<?php


namespace PhpOffice\PhpWord\Shared;


class Font
{
    
    public static function fontSizeToPixels($fontSizeInPoints = 12)
    {
        return Converter::pointToPixel($fontSizeInPoints);
    }

    
    public static function inchSizeToPixels($sizeInInch = 1)
    {
        return Converter::inchToPixel($sizeInInch);
    }

   
    public static function centimeterSizeToPixels($sizeInCm = 1)
    {
        return Converter::cmToPixel($sizeInCm);
    }

   
    public static function centimeterSizeToTwips($sizeInCm = 1)
    {
        return Converter::cmToTwip($sizeInCm);
    }

    
    public static function inchSizeToTwips($sizeInInch = 1)
    {
        return Converter::inchToTwip($sizeInInch);
    }

    
    public static function pixelSizeToTwips($sizeInPixel = 1)
    {
        return Converter::pixelToTwip($sizeInPixel);
    }

    
    public static function pointSizeToTwips($sizeInPoint = 1)
    {
        return Converter::pointToTwip($sizeInPoint);
    }
}
