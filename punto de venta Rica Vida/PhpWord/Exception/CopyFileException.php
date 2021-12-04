<?php


namespace PhpOffice\PhpWord\Exception;


final class CopyFileException extends Exception
{
    /**
     * @param string $source The fully qualified source file name.
     * @param string $destination The fully qualified destination file name.
     * @param integer $code The user defined exception code.
     * @param \Exception $previous The previous exception used for the exception chaining.
     */
    final public function __construct($source, $destination, $code = 0, \Exception $previous = null)
    {
        parent::__construct(
            sprintf('Could not copy \'%s\' file to \'%s\'.', $source, $destination),
            $code,
            $previous
        );
    }
}
