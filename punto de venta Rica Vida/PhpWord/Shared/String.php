<?php


namespace PhpOffice\PhpWord\Shared;


class String
{
    
    private static $controlCharacters = array();

    
    public static function controlCharacterOOXML2PHP($value = '')
    {
        if (empty(self::$controlCharacters)) {
            self::buildControlCharacters();
        }

        return str_replace(array_keys(self::$controlCharacters), array_values(self::$controlCharacters), $value);
    }

   
    public static function controlCharacterPHP2OOXML($value = '')
    {
        if (empty(self::$controlCharacters)) {
            self::buildControlCharacters();
        }

        return str_replace(array_values(self::$controlCharacters), array_keys(self::$controlCharacters), $value);
    }

    
    public static function isUTF8($value = '')
    {
        return $value === '' || preg_match('/^./su', $value) === 1;
    }

    
    public static function toUTF8($value = '')
    {
        if (!is_null($value) && !self::isUTF8($value)) {
            $value = utf8_encode($value);
        }

        return $value;
    }

    
    public static function toUnicode($text)
    {
        return self::unicodeToEntities(self::utf8ToUnicode($text));
    }

    
    private static function utf8ToUnicode($text)
    {
        $unicode = array();
        $values = array();
        $lookingFor = 1;

        for ($i = 0; $i < strlen($text); $i++) {
            $thisValue = ord($text[$i]);
            if ($thisValue < 128) {
                $unicode[] = $thisValue;
            } else {
                if (count($values) == 0) {
                    $lookingFor = $thisValue < 224 ? 2 : 3;
                }
                $values[] = $thisValue;
                if (count($values) == $lookingFor) {
                    if ($lookingFor == 3) {
                        $number = (($values[0] % 16) * 4096) + (($values[1] % 64) * 64) + ($values[2] % 64);
                    } else {
                        $number = (($values[0] % 32) * 64) + ($values[1] % 64);
                    }
                    $unicode[] = $number;
                    $values = array();
                    $lookingFor = 1;
                }
            }
        }

        return $unicode;
    }

    
    private static function unicodeToEntities($unicode)
    {
        $entities = '';

        foreach ($unicode as $value) {
            if ($value != 65279) {
                $entities .= $value > 127 ? '\uc0{\u' . $value . '}' : chr($value);
            }
        }

        return $entities;
    }

    
    public static function removeUnderscorePrefix($value)
    {
        if (!is_null($value)) {
            if (substr($value, 0, 1) == '_') {
                $value = substr($value, 1);
            }
        }

        return $value;
    }

   
    private static function buildControlCharacters()
    {
        for ($i = 0; $i <= 19; ++$i) {
            if ($i != 9 && $i != 10 && $i != 13) {
                $find = '_x' . sprintf('%04s', strtoupper(dechex($i))) . '_';
                $replace = chr($i);
                self::$controlCharacters[$find] = $replace;
            }
        }
    }
}
