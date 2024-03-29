eiseXLSX
========

XLSX file data read-write library that operates with native cell addresses like A1 or R1C1.

This class was designed for server-side manipulations with uploaded spreadsheets in Microsoft® Excel™ 2007-2011-2013 file format – OpenXML SpereadsheetML.

Current version of this library allows to read user-uploaded file contents and to write data to preliminary uploaded template file and send it back to the user:
- it allows to change existing cell data
- clone rows and fill-in new rows with data
- clone sheets within workbook, remove unnecessary sheets
- colorization of cells.

This library offers an easiest way to make Excel™-based EDI with PHP-based information systems, for data input and output.

Users are no longer need to convert Excel™ spreadsheets to CSV and other formats, they can simply upload data to the website using their worksheets.

You can use files received from users as your website’s output document templates with 100% match of cell formats, sheet layout, design, etc. With eiseXLSX you can stop wasting your time working on output documents layout – you can just ask your customer staff to prepare documents they’d like to see in XLSX format. Then you can put these files to the server and fill them with necessary data using PHP.

Unlike other PHP libraries for XLSX files manipulation eiseXLSX is simple, compact and laconic. You don’t need to learn XLSX file format to operate with it. Just use only cell addresses in any formats (A1 and R1C1 are supported) and data from your website database. As simple as that.

Donate on Patreon: https://www.patreon.com/easyise

Project home: <http://russysdev.github.io/eiseXLSX/>
On-line Reference Manual: <https://russysdev.github.io/eiseXLSX/docs>

PHPClasses home: <https://www.phpclasses.org/package/8573-PHP-Read-and-write-Excel-spreadsheets-in-XLSX-format.html>

Examples

Write:

- Obtain formatted XLSX from your customer
- Unzip it to the folder inside your project (or any other folder you like)
- Make PHP script and include similar code inside:

```
<?php
include_once "eiseXLSX/eiseXLSX.php";

$xlsx = new eiseXLSX("myemptyfile.xlsx");

$xlsx->data('A4', 'Hello, world!'); // either A1 or R1C1 are accepted

$xlsx->Output("mynewfile.xlsx", "F"); // save the file
?>
```

Read:

```
<?php
include_once "eiseXLSX.php";
        
try { //give it a try to avoid any uncaught error 
      // caused by broken content of uploaded file
    $xlsx = new eiseXLSX($_FILES["fileXLSX"]["tmp_name"]);
} catch(eiseXLSX_Exception $e) {
    die($e->getMessage());
}

echo ($myData = $xlsx->data("R15C10")); //voilat!
?>
```

Latest news: 

- __1.6.1.0__: Now there's the unlockSheets() method that removes necessity to enter password to modify any data/run macro by withdrawing of sheetProtection tag from all or some specified sheets. \
Example:
```
$xlsx->unlockSheets(); // unlocks all sheets
```
or
```
$xlsx->unlockSheets(1); // unlocks sheet with sheetId=1
```
or
```
$xlsx->unlockSheets([1, 2, 4]); // unlocks sheet with sheetId's 1, 2 and 4.
```


- __1.6.0.4__: I've added an ability to keep formulas in cells when data is set. Function data() is updated with parameter $flags, put there ['keep_formula'=>True] and formula will be kept for given cell. \
Example:
```
$xlsx->data('R1C1', 42, 'n', ['keep_formula' => True]);
```

- WARNING: eiseXLSX::Output() function behaviour is changed. See more at <https://russysdev.github.io/eiseXLSX/docs#eisexlsx-output>.
- method eiseXLSX::getDataValidationList($cellAddress) - returns data validation list as associative array.
- static method eiseXLSX::checkAddressInRange() - checks whether cell address belong to given range or not.
- method eiseXSLX::getDataByRange() - returns associative array of data in specified cell range.

&copy;2013-2022 Ilya S. Eliseev \
Licensed under GNU Public License v 2.0
