<?php


namespace PhpOffice\PhpWord\Exception;


final class CreateTemporaryFileException extends Exception
{
    
    final public function __construct($code = 0, \Exception $previous = null)
    {
        parent::__construct(
            'Could not create a temporary file with unique name in the specified directory.',
            $code,
            $previous
        );
    }
}
