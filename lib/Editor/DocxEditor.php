<?php

namespace Ruvents\OfficeTools\Editor;

use Ruvents\OfficeTools\Exception\ZipOpenException;

class DocxEditor
{
    const DOCUMENT_NAME = 'word/document.xml';

    /**
     * @var \ZipArchive
     */
    private $originalZip;

    /**
     * @var string
     */
    private $data;

    /**
     * @param string|\SplFileInfo $file
     *
     * @throws ZipOpenException
     */
    public function __construct($file)
    {
        $this->originalZip = new \ZipArchive();

        if (true !== $code = $this->originalZip->open($file)) {
            throw new ZipOpenException($file, $code);
        }

        $this->data = $this->originalZip->getFromName(self::DOCUMENT_NAME);
    }

    /**
     * @param array $pairs
     *
     * @return $this
     */
    public function replace(array $pairs)
    {
        $pairs = array_map(function ($value) {
            return htmlspecialchars($value, ENT_XML1);
        }, $pairs);

        $this->data = strtr($this->data, $pairs);

        return $this;
    }

    /**
     * @param string|\SplFileInfo $file
     * @param int                 $flags
     *
     * @return $this
     * @throws ZipOpenException
     */
    public function saveTo($file, $flags = \ZipArchive::CREATE)
    {
        $zip = new \ZipArchive();

        if (true !== $code = $zip->open($file, $flags)) {
            throw new ZipOpenException($file, $code);
        }

        for ($i = 0; $i < $this->originalZip->numFiles; $i++) {
            $name = $this->originalZip->getNameIndex($i);
            $contents = self::DOCUMENT_NAME === $name ? $this->data : $this->originalZip->getFromName($name);
            $zip->addFromString($name, $contents);
        }

        $zip->close();

        return $this;
    }

    public function __destruct()
    {
        $this->originalZip->close();
    }
}
