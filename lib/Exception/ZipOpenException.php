<?php

namespace Ruvents\OfficeTools\Exception;

class ZipOpenException extends RuntimeException
{
    public function __construct($file, $code)
    {
        parent::__construct(
            sprintf('ZipArchive failed to open file "%s". Error code: %d', $file, $code), $code
        );
    }
}
