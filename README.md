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

		// Output
		/*
			Array
			(
			    [RowId] => 10001
			    [Cells] => Array
			        (
			            [0]  => Cell Value 0
			            [1]  => Cell Value 1
			            [2]  => Cell Value 2
			            [3]  => Cell Value 3
			            [5]  => Cell Value 5
			            [6]  => Cell Value 6
			            [7]  => Cell Value 7
			            [8]  => Cell Value 8
			            [9]  => Cell Value 9
			            [10] => Cell Value 10
			            [11] => Cell Value 11
			            [12] => Cell Value 12
			            [13] => Cell Value 13
			            [14] => Cell Value 14
			            [15] => Cell Value 15
			            [16] => Cell Value 16
			            [17] => Cell Value 17
			            [18] => Cell Value 18
			            [21] => Cell Value 21
			        )
			)
		*/
	}
}
```
