<?php
/**
 * Author: Andre Sieverding
 * Date: 08.10.2018
 */

// Translate Excel file

// Deactivate error reporting
error_reporting(0);
ini_set('display_errors', 0);

// Ignoring user abort, so the script can be executed completly
ignore_user_abort(true);

// Deactivate timeout
set_time_limit(0);

// Set higher memory limit
ini_set('memory_limit', '1024M');

// Define DeepL Pro API Key
define('API_KEY', '123');

// Define rows for language code and text value
define('ROW_LANG', 1);
define('ROW_VAL', 3);

// Define excel filename for translating
define('EXC_FILE', 'Merkmale1.xlsx');

// Function for translating using DeepL Pro API translator
function translate ($text, $target_lang) {
	$apiLink = "https://api.deepl.com/v1/translate?auth_key=" . API_KEY . "&text=" . urlencode($text) . "&source_lang=DE&target_lang=" . $target_lang;
	$apiCallback = file_get_contents($apiLink);

	if ($apiCallback != '') {
		$apiObject = json_decode($apiCallback);
		$target = $apiObject->translations[0]->text;

		return $target;
	} else {
		return false;
	}
}

// Include composer autoloader
require('./vendor/autoload.php');

// Use PhpSpreadsheet classes
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Reader;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as Writer;

// Initialize new excel reader
$reader = new Reader();

// Load our excel file for translating
$spreadsheet = $reader->load("./src/" . EXC_FILE);

// Get worksheets of this excel file
$worksheets = $spreadsheet->getAllSheets();

// Define some useful vars
$rawData = array();
$finalData = array();
$n = 0;
$ignoredTexts = array();

// Iterate through first worksheet an get rows / cell-values
foreach ($worksheets[0]->getRowIterator() as $row) {
	// Ignore headline
	if ($n != 0) {
		// Get cells
		$cellIterator = $row->getCellIterator();
		$cellIterator->setIterateOnlyExistingCells(true);
		$cells = [];

		foreach ($cellIterator as $cell) {
			// Read cell value
			$cells[] = $cell->getValue();
		}

		$cells[ROW_LANG] = strtoupper(trim($cells[ROW_LANG]));

		// Notice already translated texts by saving their text-id
		if ($cells[ROW_LANG] != 'DE') {
			$ignoredTexts[$cells[0]][] = $cells[ROW_LANG];
		}

		// Add row of data to raw data array
		$rawData[] = $cells;
	}

	$n++;
}

// Translate all not translated texts using Deepl Pro API translation service
for ($i = 0, $j = count($rawData); $i < $j; $i++) {
	$alreadyTranslatedLangKeys = array();
	
	// Add data to final data
	$finalData[] = $rawData[$i];

	// If the current text is in German, then translate...
	if ($rawData[$i][ROW_LANG] == 'DE') {
		if (array_key_exists($rawData[$i][0], $ignoredTexts)) {
			$alreadyTranslatedLangKeys = $ignoredTexts[$rawData[$i][0]];
		}

		// ...it into Englisch, if the translation doesn't already exists
		if (!in_array('EN', $alreadyTranslatedLangKeys)) {
			$enRawData = $rawData[$i];
			$enRawData[ROW_LANG] = "EN";
			$enRawData[ROW_VAL] = translate($rawData[$i][ROW_VAL], 'EN');

			if ($enRawData[ROW_VAL] !== false) {
				$finalData[] = $enRawData;
			}
		}

		// ...it into France, if the translation doesn't already exists
		if (!in_array('FR', $alreadyTranslatedLangKeys)) {
			$frRawData = $rawData[$i];
			$frRawData[ROW_LANG] = "FR";
			$frRawData[ROW_VAL] = translate($rawData[$i][ROW_VAL], 'FR');

			if ($enRawData[ROW_VAL] !== false) {
				$finalData[] = $frRawData;
			}
		}
	}
}

// Iterate through final data an write data into second worksheet
for ($i = 0, $j = count($finalData); $i < $j; $i++) {
	$worksheets[1]->setCellValue('A' . ($i + 2), $finalData[$i][0]);
	$worksheets[1]->setCellValue('B' . ($i + 2), $finalData[$i][1]);
	$worksheets[1]->setCellValue('C' . ($i + 2), $finalData[$i][2]);
	$worksheets[1]->setCellValue('D' . ($i + 2), $finalData[$i][3]);
	$worksheets[1]->setCellValue('E' . ($i + 2), $finalData[$i][4]);
	$worksheets[1]->setCellValue('F' . ($i + 2), $finalData[$i][5]);
	$worksheets[1]->setCellValue('G' . ($i + 2), $finalData[$i][6]);
	$worksheets[1]->setCellValue('H' . ($i + 2), $finalData[$i][7]);
	$worksheets[1]->setCellValue('I' . ($i + 2), $finalData[$i][8]);
	$worksheets[1]->setCellValue('J' . ($i + 2), $finalData[$i][9]);
}

// Save excel file
$writer = new Writer($spreadsheet);
$writer->save('./target/' . EXC_FILE);

echo 'Ready!';
