# RUVENTS PHP Office Tools

## Installation

`$ composer require ruvents/office-tools`

## Usage

### .docx editor

```php
<?php

use Ruvents\OfficeTools\Editor\DocxEditor;

(new DocxEditor('original.docx'))
    ->replace([
        '__first_name__' => 'Ivan',
        '__last_name__' => 'Ivanov',
    ])
    ->saveTo('target.docx', \ZipArchive::OVERWRITE);
```
