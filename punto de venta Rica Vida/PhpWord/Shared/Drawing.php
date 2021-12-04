<?php


namespace PhpOffice\PhpWord\Shared;


class Drawing
{
    public static function pixelsToEMU($pValue = 0)
    {
        return round($pValue * 9525);
    }

   
    public static function emuToPixels($pValue = 0)
    {
        if ($pValue != 0) {
            return round($pValue / 9525);
        } else {
            return 0;
        }
    }

    
    public static function pixelsToPoints($pValue = 0)
    {
        return $pValue * 0.67777777;
    }

    
    public static function pointsToPixels($pValue = 0)
    {
        if ($pValue != 0) {
            return $pValue * 1.333333333;
        } else {
            return 0;
        }
    }

    public static function degreesToAngle($pValue = 0)
    {
        return (integer)round($pValue * 60000);
    }

   
    public static function angleToDegrees($pValue = 0)
    {
        if ($pValue != 0) {
            return round($pValue / 60000);
        } else {
            return 0;
        }
    }

   
    public static function pixelsToCentimeters($pValue = 0)
    {
        return $pValue * 0.028;
    }

   
    public static function centimetersToPixels($pValue = 0)
    {
        if ($pValue != 0) {
            return $pValue / 0.028;
        } else {
            return 0;
        }
    }

    
    public static function centimetersToTwips($pValue = 0)
    {
        if ($pValue != 0) {
            return $pValue * 566.928;
        } else {
            return 0;
        }
    }

   
    public static function twipsToCentimeters($pValue = 0)
    {
        if ($pValue != 0) {
            return $pValue / 566.928;
        } else {
            return 0;
        }
    }

   
    public static function inchesToTwips($pValue = 0)
    {
        if ($pValue != 0) {
            return $pValue * 1440;
        } else {
            return 0;
        }
    }

    
    public static function twipsToInches($pValue = 0)
    {
        if ($pValue != 0) {
            return $pValue / 1440;
        } else {
            return 0;
        }
    }

    
    public static function twipsToPixels($pValue = 0)
    {
        if ($pValue != 0) {
            return round($pValue / 15.873984);
        } else {
            return 0;
        }
    }

    
    public static function htmlToRGB($pValue)
    {
        if ($pValue[0] == '#') {
            $pValue = substr($pValue, 1);
        }

        if (strlen($pValue) == 6) {
            list($colorR, $colorG, $colorB) = array($pValue[0] . $pValue[1], $pValue[2] . $pValue[3], $pValue[4] . $pValue[5]);
        } elseif (strlen($pValue) == 3) {
            list($colorR, $colorG, $colorB) = array($pValue[0] . $pValue[0], $pValue[1] . $pValue[1], $pValue[2] . $pValue[2]);
        } else {
            return false;
        }

        $colorR = hexdec($colorR);
        $colorG = hexdec($colorG);
        $colorB = hexdec($colorB);

        return array($colorR, $colorG, $colorB);
    }
}
