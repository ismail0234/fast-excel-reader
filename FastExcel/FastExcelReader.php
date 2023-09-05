<?php 

namespace FastExcel;

class FastExcelReader
{
	/**
	 *
	 * workbookSharedStringName yolunu barındırır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private $workbookSharedStringName = 'xl/sharedStrings.xml';

	/**
	 *
	 * workbookSheet1Name yolunu barındırır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private $workbookSheet1Name = 'xl/worksheets/sheet1.xml';

	/**
	 *
	 * Excel yolunu barındırır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private $FilePath = null;

	/**
	 *
	 * TmpPath yolunu barındırır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private $TmpPath = null;

	/**
	 *
	 * TmpPathId yolunu barındırır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private $TmpPathId = null;

	/**
	 *
	 * Zip Nesnesini barındırır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private $Zip = null;

	/**
	 *
	 * SharedStrings Nesnesini barındırır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private $SharedStrings = array();

	/**
	 *
	 * Aralık uzunluluğunu barındırır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private $seekLength = 1024;

	/**
	 *
	 * Mevcut konum uzunluluğunu barındırır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private $currentSeekPosition = 0;

	/**
	 *
	 * Excel kolon ad ve indexlerini barındırır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private $excelColumnIndexes = array();

	/**
	 *
	 * Sınıf ayarlamalarını yapar.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	public function __construct()
	{
		$this->setSeekLength(128);
		$this->genereateColumnIndexes();
	}

	/**
	 *
	 * Kolon indekslerini oluşturur.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	public function genereateColumnIndexes()
	{
		$letter = 'A';
		$index  = 0;

		while ($letter !== 'ZZ') 
		{
		    $this->excelColumnIndexes[$letter++] = $index++;
		}
	}

	/**
	 *
	 * Aralık uzunluğunu değiştirir.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	public function setSeekLength($length)
	{
		$this->seekLength = 1024 * $length;
	}

	/**
	 *
	 * Excel yolunu ayarlar.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	public function SetPath($filePath)
	{
		$this->FilePath = trim($filePath);
	}

	/**
	 *
	 * Tmp yolunu ayarlar.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	public function SetTmp($tmpPath)
	{
		$this->TmpPath = trim($tmpPath);

		if (substr($this->TmpPath, -1) != '/')
		{
			$this->TmpPath .= '/';
		}
	}

	/**
	 *
	 * Tmp yolunu ayarlar.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	public function GetRandomTempPath($filePath = null)
	{
		if ($filePath)
		{
			return sprintf('%s/%s', $this->GetRandomTempPath(null), $filePath);
		}

		return sprintf('%s%s', $this->TmpPath, $this->TmpPathId);
	}

	/**
	 *
	 * Dosyayı açar ve işlemleri başlatır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	public function Open()
	{
		if (empty($this->FilePath) || empty($this->TmpPath) || !file_exists($this->FilePath))
		{
			return false;
		}

		if (!file_exists($this->TmpPath) && !mkdir($this->TmpPath, 0777)) 
		{
			return false;
		}

		$this->currentSeekPosition = 0;
		$this->SharedStrings       = array();
		$this->TmpPathId           = sprintf('%s_%s', uniqid(), rand(1000000, 9000000));
		$this->Zip                 = new \ZipArchive();

		if ($this->Zip->open($this->FilePath) === true)
		{
			$this->Zip->extractTo($this->GetRandomTempPath(), $this->workbookSharedStringName);
			$this->Zip->extractTo($this->GetRandomTempPath(), $this->workbookSheet1Name);

			$this->loadSharedStrings();
			return true;
		}

		return false;
	}

	/**
	 *
	 * Satırları döner.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	public function GetRows()
	{
		$stream = fopen($this->GetRandomTempPath($this->workbookSheet1Name), 'r');
		$parser = xml_parser_create();

		while (($data = fread($stream, $this->seekLength))) 
		{
			xml_parse($parser, $data);

			foreach ($this->getRowDetails($data) as $row)
			{
				yield $row;
			}

			if (strlen($data) < $this->seekLength)
			{
				break;
			}

			$this->currentSeekPosition += $this->getSeekPosition($data) + 6;

			fseek($stream, $this->currentSeekPosition);
		}

		xml_parser_free($parser);

		fclose($stream);

	}

	/**
	 *
	 * Kütüphane yok edilirken tetiklenir.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	public function __destruct()
	{
		if ($this->Zip) 
		{
			$this->Zip->close();
		}

		if (!empty($this->GetRandomTempPath()) && file_exists($this->GetRandomTempPath())) 
		{
			$this->removeDirectory($this->GetRandomTempPath());
		}
	}

	/**
	 *
	 * Satır detaylarını döner.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private function getRowDetails($data)
	{
		$rows = [];

		if (preg_match_all('@<row r="(.*?)"(.*?)>(.*?)</row>@', $data, $rowMatch))
		{
			foreach ($rowMatch[3] as $rowIndex => $row)
			{
				$rows[] = array(
					'RowId' => $rowMatch[1][$rowIndex],
					'Cells' => $this->getCellDetails($rowMatch[1][$rowIndex], $rowMatch[3][$rowIndex]),
				);
			}
		}

		return $rows;
	}

	/**
	 *
	 * Sütun detaylarını döner.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private function getCellDetails($rowId, $cellData)
	{
		$cells = [];

		if (preg_match_all('@<c r="([A-Z\/]+)(.*?)"(.*?)><v>(.*?)</v></c>@', $cellData, $cellMatch))
		{
			foreach ($cellMatch[4] as $cellIndex => $cell)
			{
				$cells[$this->excelColumnIndexes[$cellMatch[1][$cellIndex]]] = $this->getCellValue($cell, $cellMatch[3][$cellIndex]);
			}
		}

		return $cells;
	}

	/**
	 *
	 * Hücre değerini döner.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private function getCellValue($value, $cellProperty)
	{
		if (strpos($cellProperty, "t=\"s\"") !== false)
		{
			return $this->SharedStrings[$value];
		}

		return $value;

	}

	/**
	 *
	 * Aralık konumunu döner.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private function getSeekPosition($data)
	{
		$seek = 0;

		do 
		{
			$seek += 300;

			$position = strpos($data, '</row>', $this->seekLength - $seek);
		} 
		while ($position === false);

		return $position;
	}

	/**
	 *
	 * Paylaşılan metinleri önbelleğe alır.
	 *
	 * @author Ismail <ismaiil_0234@hotmail.com>
	 *
	 */
	private function loadSharedStrings()
	{
		$xml = new \XMLReader();
		$xml->open($this->GetRandomTempPath($this->workbookSharedStringName));

		$currentIndex = -1;

		while ($xml->read()) 
		{
			if ($xml->nodeType === \XMLReader::ELEMENT) 
			{
				switch ($xml->name) 
				{
					case 'si':
						$currentIndex++;
						break;
					case 't':
						$this->SharedStrings[$currentIndex] = $xml->readString();
						break;
				}
			}
		}

		$xml->close();
	}

	/**
	 *
	 * Klasör ve bağlı alt klasörleri/dosyaları siler.
	 *
	 * @author Ismail Satilmis <ismaiil_0234@hotmail.com>
	 *
	 */
	private function removeDirectory($directory) 
	{ 
		if (is_dir($directory)) 
		{ 
			$files = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($directory, \RecursiveDirectoryIterator::SKIP_DOTS), \RecursiveIteratorIterator::CHILD_FIRST);

			foreach ($files as $finfo) 
			{
				if ($finfo->isDir())
				{
					rmdir($finfo->getRealPath());
				}
				else
				{
					unlink($finfo->getRealPath());
				}
			}

			rmdir($directory);
		} 
	}
}
