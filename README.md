[![Latest Stable Version](https://poser.pugx.org/ismail0234/fast-excel-reader/v/stable)](https://packagist.org/packages/ismail0234/fast-excel-reader)
[![Total Downloads](https://poser.pugx.org/ismail0234/fast-excel-reader/downloads)](https://packagist.org/packages/ismail0234/fast-excel-reader)
[![License](https://poser.pugx.org/ismail0234/fast-excel-reader/license)](https://packagist.org/packages/ismail0234/fast-excel-reader)

A very fast excel reader written in php.

## Composer Install

```php

composer require ismail0234/fast-excel-reader

```

## Benchmark

| Information        | 100.000 Row + 20 Column | 10.000 Row + 20 Column |
|--------------------|------------------------|-------------------------|
| PhpSpreadsheet     | 120.50 Seconds         | 11.52 Seconds           |
| Box/Spout Excel    | 101.45 Seconds         | 12.56 Seconds           |
| PHPExcel           | 72.06 Seconds          | 6.64 Seconds            |
| Akeneo Spreadsheet | 17.94 Seconds          | 2.06 Seconds            |
| FastExcelReader    | 3.97 Seconds           | 0.44 Seconds            |

## Install & Usage

```php

use FastExcel\FastExcelReader;

include "vendor/autoload.php";

$excel = new FastExcelReader();
$excel->SetPath('/home/mywebsite/public_html/myexcel.xlsx');
$excel->SetTmp('/home/mywebsite/public_html/tmp/');

if ($excel->Open())
{
	foreach ($excel->GetRows() as $row) 
	{
		// Row Details
		echo '<pre>';
		print_r($row);
		echo '</pre>';
		exit;
	}
}
```
