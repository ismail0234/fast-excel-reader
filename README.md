[![Latest Stable Version](https://poser.pugx.org/ismail0234/fast-excel-reader/v/stable)](https://packagist.org/packages/ismail0234/fast-excel-reader)
[![Total Downloads](https://poser.pugx.org/ismail0234/fast-excel-reader/downloads)](https://packagist.org/packages/ismail0234/fast-excel-reader)
[![License](https://poser.pugx.org/ismail0234/fast-excel-reader/license)](https://packagist.org/packages/ismail0234/fast-excel-reader)

A very fast excel reader written in php.

## Composer Install

```php

composer require ismail0234/fast-excel-reader

```


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
