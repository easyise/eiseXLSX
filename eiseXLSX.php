<?php
/**
 * eiseXLSX
 * ===
 * 
 * XLSX file data read-write library that operates with native cell addresses like A1 or R1C1.
 *
 * This class was designed for server-side manipulations with uploaded spreadsheets in Microsoft® Excel™ 2007-2011-2013 file format – OpenXML SpereadsheetML.
 *
 * Current version of this library allows to read user-uploaded file contents and to write data to preliminary uploaded template file and send it back to the user:
 * * it allows to change existing cell data
 * * clone rows and fill-in new rows with data
 * * clone sheets within workbook, remove unnecessary sheets
 * * colorization of cells.
 *
 * This library offers the easiest way to make Excel™-based data exchange with information systems written in PHP.
 * 
 * Users are no longer need to convert Excel™ spreadsheets to CSV and other formats, they can simply upload data to the website using their worksheets.
 * 
 * You can use files received from users as your website’s output document templates with 100% match of cell formats, sheet layout, design, etc. With eiseXLSX you can stop wasting your time working on output documents layout – you can just ask your customer staff to prepare documents they’d like to see in XLSX format. Then you can put these files to the server and fill them with necessary data using PHP.
 * 
 * Unlike other PHP libraries for XLSX files manipulation eiseXLSX is simple, compact and laconic. You don’t need to learn XLSX file format to operate with it. Just use only cell addresses in any formats (A1 and R1C1 are supported) and data from your website database. As simple as that.
 * 
 * Project home: <http://russysdev.github.io/eiseXLSX/>
 * On-line Reference Manual: <https://russysdev.github.io/eiseXLSX/docs>
 *   
 * >  Based on:  
 * >  Simple XLSX   
 * >  @author kirik <mail@kirik.ws>  
 * >  @version 0.1  
 * >  Developed under GNU General Public License, version 3: <http://www.gnu.org/licenses/lgpl.txt>
 *
 * @uses SimpleXML, DOM
 *
 * @package eiseXLSX <https://github.com/easyise/eiseXLSX>
 *   
 * @author Ilya Eliseev <ie@e-ise.com>, Continutors: Igor Zhuravlev, Dmitry Zakharov
 * @copyright (c) 2012-2017 Ilya S. Eliseev
 *
 * @license GNU Public License <http://opensource.org/licenses/gpl-license.php>
 *
 * @version 1.9beta
 *
 */

/**
 * The class that creates objects with Excel workbooks inside. Public class methods are allowing to make any declared manupulations with the workbook.
 */
class eiseXLSX {

/** @ignore */
const DS = DIRECTORY_SEPARATOR;
/** @ignore */
const Date_Bias = 25569; // number of days between Excel and UNIX epoch
/** @ignore */
const VERSION = '1.6';
/** @ignore */
const TPL_DIR = 'templates';

/** @ignore */
private $_parsed = array();
/** @ignore */
private $arrXMLs = array(); // all XML files
/** @ignore */
private $arrSheets = array(); // all sheets
/** @ignore */
private $arrSheetPath = array(); // all paths to sheets
/** @ignore */
private $_cSheet; // current sheet

/** 
 * Default file name for output
 */
/** @ignore */
public $defaultFileName = 'eiseXLSX.xlsx';

/**
 * Indexed colors acсoring to XLSX file standard
 */
/** @ignore */
static $arrIndexedColors = Array('00000000', '00FFFFFF', '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', '0000FFFF', '00000000', '00FFFFFF', '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', '0000FFFF', '00800000', '00008000', '00000080', '00808000', '00800080', '00008080', '00C0C0C0', '00808080', '009999FF', '00993366', '00FFFFCC', '00CCFFFF', '00660066', '00FF8080', '000066CC', '00CCCCFF', '00000080', '00FF00FF', '00FFFF00', '0000FFFF', '00800080', '00800000', '00008080', '000000FF', '0000CCFF', '00CCFFFF', '00CCFFCC', '00FFFF99', '0099CCFF', '00FF99CC', '00CC99FF', '00FFCC99', '003366FF', '0033CCCC', '0099CC00', '00FFCC00', '00FF9900', '00FF6600', '00666699', '00969696', '00003366', '00339966', '00003300', '00333300', '00993300', '00993366', '00333399', '00333333');


/**
 * The constructor. It reads Excel workbook supplied as the template or data source to read. It can be either XLSX file or unzipped one, into the directory speicfied in $templatePath parameter. Unzipping XLSX files will reduce your server CPU load ;).
 *
 * Also it parses all XMLs inside the workbook and makes all the necessary preparations for future data read and manipulations.
 * 
 * If $templatePath parameter is omitted it reads default template. If path's not found it throws an exception (object of eiseXLSX_Exception class).
 *
 * @category Read / Write
 * 
 * @param string $templatePath The path to Excel workbook file or directory.
 */
public function __construct( $templatePath='' ) {

    if(!$templatePath){
        $templatePath = 'empty';
    } else {
        $this->defaultFileName = basename($templatePath);
    }

    // read template
    $templatePath = (file_exists($templatePath) 
        ?  $templatePath 
        :  dirname( __FILE__ ).self::DS. self::TPL_DIR .self::DS.$templatePath
    );
    
    if (!file_exists($templatePath)){
    
        throw new eiseXLSX_Exception("XLSX template ({$templatePath}) does not exist");
    
    }
    
    if (is_dir($templatePath)){
        $eiseXLSX_FS = new eiseXLSX_FS($templatePath);
        list($arrDir, $arrFiles) = $eiseXLSX_FS->get();
    } else {
        list($arrDir, $arrFiles) = $this->unzipToMemory($templatePath);
    }
    
    $nSheets = 0; $nFirstSheet = 1;
    foreach($arrFiles as $path => $contents) {
        
        $path = "/".str_replace(self::DS, "/", $path);
        
        $this->arrXMLs[$path] = @simplexml_load_string($contents);
        
        if (empty($this->arrXMLs[$path])){
            $this->arrXMLs[$path] = (string)$contents;
        }
        
        if (preg_match("/\.rels$/", $path)){
            foreach($this->arrXMLs[$path]->Relationship as $Relationship){
                if((string)$Relationship["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument") {
                    $this->officeDocumentPath = self::getPathByRelTarget($path, (string)$Relationship["Target"]);
                    $this->officeDocument = &$this->arrXMLs[$this->officeDocumentPath];
                }
                
                if((string)$Relationship["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings") {
                    $this->sharedStrings = &$this->arrXMLs[self::getPathByRelTarget($path, (string)$Relationship["Target"])];
                }
                
                if((string)$Relationship["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
                    $this->theme = &$this->arrXMLs[self::getPathByRelTarget($path, (string)$Relationship["Target"])];
                }
                
                if((string)$Relationship["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles") {
                    $this->styles = &$this->arrXMLs[self::getPathByRelTarget($path, (string)$Relationship["Target"])];
                }
                
            }
        }
        
    }
    
    $this->officeDocumentRelPath = self::getRelFilePath($this->officeDocumentPath);
    $ix = 0;
    foreach($this->officeDocument->sheets->sheet as $sheet) {
        //<sheet r:id="rId1" sheetId="1" name="ACT"/>
        $relId = $sheet->attributes('r', TRUE)->id;
        
        foreach($this->arrXMLs[$this->officeDocumentRelPath]->Relationship as $Relationship){
            if ((string)$Relationship["Id"]==$relId){
                $path = self::getPathByRelTarget($this->officeDocumentRelPath, (string)$Relationship["Target"]);
                break;
            }
        }
        $this->arrSheets[(string)$sheet["sheetId"]] = &$this->arrXMLs[$path];
        $this->arrSheetPath[(string)$sheet["sheetId"]] = $path;
        $nFirstSheet = ($ix==0 ? (string)$sheet["sheetId"] : $nFirstSheet);
        $ix++;
    }
    
    $this->selectSheet($nFirstSheet);
    
}

/**
 * eiseXLSX::data() function reads or sets data for cell with specified $cellAddress. If parameter $data is omitted, function just returns current cell data. If $data contains something, function tries to set it.
 * Data types note:
 * - strings are to be returned and assigned as strings
 * - numeric values are to be returned and set as strings with numeric values in latin1 locale inside.
 * - date/time values are to be returned and set as strings formatted as 'YYYY-MM-DD HH:MM:SS'
 * 
 * @param string $cellAddress - both R1C1 and A1 address formats are acceptable. Case-insensitive. Examples: "AI75", "r10c25". 
 * @param mixed $data - data to set. If not set at function call, function just returns data. If set, function sets this data for given cell.
 * @param string $t - if omitted eiseXLSX accepts the data as string and put contents to sharedStrings.xml. Otherwise it tries to re-format date as seconds or number as real one with period as decimal separator. 
 * Possible values: 'n' - for numeric values like integer or real numbers;
 *   's' (default) - for strings, but if string can be evaluated as number using is_numeric() PHP function, numeric value will be set;
 *   'd' - for datetime values.
 *
 * @return string - cell data before new value is set (if any).
 *
 * @category Sheet manipulations
 */
public function data($cellAddress, $data = null, $t = "s"){
    
    $retVal = null;
    
    list( $x, $y, $addrA1, $addrR1C1 ) = self::cellAddress($cellAddress);
    
    $c = $this->locateCell($x, $y);
    if (!$c && $data !== null){
        $c = &$this->addCell($x, $y);
    }
    
    
    
    if (isset($c->v[0])){ // if it has value
        
        $o_v = &$c->v[0];
        if ($c["t"]=="s"){ // if existing type is string
            $siIndex = (int)$c->v[0];
            $o_si = &$this->sharedStrings->si[$siIndex];
            $retVal = strip_tags($o_si->asXML()); //return plain string without formatting
        } else { // if not or undefined
            $retVal = $this->formatDataRead($c["s"], (string)$o_v);
            if ($data!==null && $t=="s") {// if forthcoming type is string, we add shared string
                $o_si = &$this->addSharedString($c);
            }
        }
    } else {
        $retVal = null;
        if ($data!==null &&
            !(!is_object($data) && (string)$data=="")
            ) // if we'd like to set data and not to empty this cell
            if ($t=="s") {// if forthcoming type is string, we add shared string
                $o_si = &$this->addSharedString($c);
                $o_v = &$c->v[0];
            } else { // if not, value is inside '<v>' tag
                $c->addChild("v", $data);
                $o_v = &$c->v[0];
            }
    }
    
    
    if ($data!==null){ // if we set data
        
        if (!is_object($data) && (string)$data=="") { // if there's an empty string, we demolite existing data
            unset($c["t"]);
            unset($c->v[0]);
        } else { // we set received value
            unset($c->f[0]); // remove forumla
            if( is_numeric($data) && func_num_args()==2 ) // if default
                $t = "n";
            switch($t){
                case "s":
                    $this->updateSharedString($o_si, $data);
                    break;
                default:
                    $this->formatDataWrite($t, $data, $c);
                    break;
            }
        }
    }
    
    return $retVal;
    
}

/**
 * This function returns contents of drop-down list for given cell, if Data Validation / List option is activated for given cell. If there's no list, this function returns NULL, if reference to drop-down list cell range is broken, it returns FALSE. Otherwise it returns associative array with origin cell addresses as keys and cell data as values. Function eiseXLSX::getDataByRange() (that uses eiseXLSX::data()) is used to obtain cell data.
 * eiseXLSX::getDataValidatioList() can be useful when you need to obtain contents of reference tables of spreadsheet fields when you try to import the workbook into the database. 
 * NOTE: This function supports only local cell range references, within current workbook. If requested cell takes drop-down list values from another workbook, function returns FALSE.
 * NOTE: Function supports only single-row or single-columns references to drop-down cell range. Otherwise it returns FALSE.
 *
 * @param string $cellAddress - Cell address. Both R1C1 and A1 address formats are acceptable. Case-insensitive. Examples: "AI75", "r10c25".
 *
 * @return mixed - NULL if there's no data validation, associative array of drop-down values with origin cell addresses as keys and FALSE in case of broken/invalid reference to drop-down cell range.
 *
 * @category Sheet manipulations
 */
public function getDataValidationList($cellAddress){

    if($this->_cSheet->dataValidations->dataValidation)
        foreach($this->_cSheet->dataValidations->dataValidation as $ix=>$val){
            if($val['type']!='list')
                continue;
            $range = $val['sqref'];
            if( self::checkAddressInRange($cellAddress, $range) ){
                $ref = (string)$val->formula1[0];
                break;
            }
        }


    if(!$ref)
        foreach($this->_cSheet->extLst->ext as $ext){
            if($ext["uri"]!='{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}')
                continue;

            //determining xl-version related x tag
            $arrNS = $ext->getNamespaces(true);
            foreach ($arrNS as $prfx => $uri) {
                if(preg_match('/^x[0-9]*$/', $prfx)){
                    $nsX = $uri;
                    break;
                }
            }

            $chdn = $ext->children($nsX);

            foreach($chdn->dataValidations->dataValidation as $ix=>$val){

                $range = (string)$val->children('xm', true);
                if( self::checkAddressInRange($cellAddress, $range) ){
                    $ref = (string)$val->formula1[0]->children('xm', true);
                    break;
                }

            }

            if($ref)
                break;

        } 


    return ($ref ? $this->getDataByRange($ref) : null);

}

/**
 * This function returns an array of data obtained from the specified $range. This range can be as well as formula-formatted (e.g. "Sheet 2!$A1:$B12") as normal particular range (like "B15:B50"). Cell list, range list and other range formats are NOT SUPPORTED (YET).
 * Reference sheets (if any) should exist in the same workbook as current sheet.
 * Empty values are not returned. 
 * If range cannot be located, function returns FALSE.
 * 
 * @param string $range - cell range in normal format (like "A14:X14") or formula-based refrence ("Sheet 3!$Z15:$Y17").
 * 
 * @return array of data obtained from range with R1C1 address as keys and values as they've been obtained with data() function. If range cannot be located, function returns FALSE.
 *
 * @category Sheet manipulations
 */
public function getDataByRange($range){

    $arrRet = array();
    $diffSheetName = $refSheetID = '';

    $range = preg_replace('/\$([a-z0-9]+)/i', '$1', $range);

    $arrRef = explode('!', $range);

    $range = $arrRef[count($arrRef)-1];

    if($diffSheetName = (count($arrRef)>1 ? $arrRef[0] : '')){
        if( !($refSheetID = $this->findSheetByName($diffSheetName)) )
            return false;

        foreach($this->arrSheets as $id=>$sheet)
            if($sheet===$this->_cSheet){
                $curSheetID = $id;
                break;
            }

        $this->selectSheet($refSheetID);

    }

    try {
        list($aX, $aY) = self::getRangeArea($range);
    } catch (eiseXLSX_Exception $e){
        return false;
    }

    for($x = $aX[0]; $x<=$aX[1]; $x++)
        for($y = $aY[0]; $y<=$aY[1]; $y++){
            $addr = "R{$y}C{$x}";
            $dt =  $this->data($addr);
            if($dt)
                $arrRet[$addr] = $dt;
        }

    if($diffSheetName)
        $this->selectSheet($curSheetID);

    return $arrRet;

}


/**
 * checkAddressInRange() function checks whether given cell belong to specified cell address range.
 *
 * @param string $adrNeedle - cell address to check. Both R1C1 and A1 address formats are acceptable. Case-insensitive. Examples: "AI75", "r10c25".
 * @param string $adrHaystack - cell address range. Both R1C1 and A1 address formats are acceptable. Can be as single cell, cell range (cell1:cell2) and list of cells and ranges, space-separated. Case-insensitive. Examples: "AI75:AJ86", "r10c25:r1c25 ", "C168 AF113:AG116 AI113:AI116 L113:N116".
 *
 * @return boolean - true if cell belongs to the range, false otherwise
 *
 * @category Cell address routines
 */
public static function checkAddressInRange($adrNeedle, $adrHaystack){

    list($xNeedle, $yNeedle) = self::cellAddress($adrNeedle);

    $arrHaystack = explode(' ', $adrHaystack);
    foreach($arrHaystack as $range){
        
        list($x, $y) = self::getRangeArea($range);

        if($x[0]<=$xNeedle && $xNeedle<=$x[1]
            && $y[0]<=$yNeedle && $yNeedle<=$y[1]){
            return true;
        }

    }   

    return false;

}

/**
 * This function returns array of top-left and bottom-right coordinates of particular range area.
 * 
 * @param $range string - cell address range. Both R1C1 and A1 address formats are acceptable. Can be as single cell or cell range (cell1:cell2). Case-insensitive. Examples: "AI75:AJ86", "r10c25:r1c25".
 *
 * @return array - array(array($x_left, $x_right), array($y_top, $y_bottom)) where x and y are column and row number correspondingly.
 *
 * @category Cell address routines
 */
public static function getRangeArea($range){

    $arrRng = explode(':', $range);
    
    list($x[0], $y[0]) = self::cellAddress($arrRng[0]);

    if($arrRng[1]) { list($x[1], $y[1]) = self::cellAddress($arrRng[1]); }
    else { $x[1] = $x[0]; $y[1] = $y[0];      }

    sort($x); sort($y);

    return array($x, $y);
}

/**
 * This method returns number of rows in active sheet.
 *
 * @return int - row number of the last row.
 *
 * @category Sheet manipulations
 */
public function getRowCount(){
    $lastRowIndex = 1;
    foreach($this->_cSheet->sheetData->row as $row){
        $lastRowIndex = (int)$row['r'];
    }
    //return count($this->_cSheet->sheetData->row);
    return $lastRowIndex;
}

/**
 * Fills cell at $cellAddress with color $fillColor or clears cell off any background color, if $fillColor is set to NULL, 0 or ''.
 * If cell is not found or color string is wrongly specified, it throws an exception.
 *
 * @param string $cellAddress Cell address, both A1 and R1C1 address formats are acceptable.
 * @param string $fillColor HTML-style color in Hex pairs, for example: #FFCC66. Should always start with hash.
 *
 * @return simpleXML object that represents specified cell.
 *
 * @category Cell decoration
 */
public function fill($cellAddress, $fillColor){
    
    $fillColor = ($fillColor ? self::colorW3C2Excel($fillColor) : "");
    
    // locate cell, if no cell - throw exception
    list( $x, $y, $addrA1, $addrR1C1 ) = self::cellAddress($cellAddress);
    $c = &$this->locateCell($x, $y);
    
    if ($c===null){
        throw new eiseXLSX_Exception('cannot apply fill - no cell at '.$cellAddress);
    }
    
    if ($fillColor){
        // locate fill by color, if no fill - add 
        $ix = 0;
        foreach($this->styles->fills->fill as $fill){
            if (strtoupper((string)$fill->patternFill->fgColor["rgb"])==$fillColor){
                $fillIx = $ix;
                break;
            }
            $ix++;
        }
        if (!isset($fillIx)){
            $xmlFill = simplexml_load_string("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{$fillColor}\"/><bgColor indexed=\"64\"/></patternFill></fill>");
            $this->insertElementByPosition((int)$this->styles->fills["count"], $xmlFill, $this->styles->fills);
            $fillIx = (int)$this->styles->fills["count"];
            $this->styles->fills["count"] = (int)$this->styles->fills["count"]+1;
        }
    } else 
        $fillIx = 0; //http://openxmldeveloper.org/discussions/formats/f/14/p/716/3685.aspx : 
        //Fill ID zero ALWAYS has to be a pattern fill (gray125). Custom fills start at index 1 and up.
    
    
    // locate style, if no style - add
    if ($c["s"]){
        $cellXf = $this->styles->cellXfs->xf[(int)$c["s"]];
        if ((int)$cellXf["fillId"] != $fillIx){ // if style is getting changed, we try to locate changed one, if we fail ,we add
            $ix = 0;
            foreach($this->styles->cellXfs->xf as $xf ){
                if ((string)$xf["borderId"]==(string)$cellXf["borderId"] && (string)$xf["fillId"]==$fillIx 
                    && (string)$xf["fontId"]==(string)$cellXf["fontId"] && (string)$xf["numFmtId"]==(string)$cellXf["numFmtId"] 
                    && (string)$xf["xfId"]==(string)$cellXf["xfId"] && (string)$xf["applyFill"]==(string)$cellXf["applyFill"] 
                    ){
                        $styleIx = $ix; break;
                    } $ix++;
            }
            if (isset($styleIx))
                $c["s"] = $styleIx;
            else {
                $xmlXF = simplexml_load_string($this->styles->cellXfs->xf[(int)$c["s"]]->asXML());
                $xmlXF["fillId"]=$fillIx; $xmlXF["applyFill"]="1";
                $this->insertElementByPosition((int)$this->styles->cellXfs["count"], $xmlXF, $this->styles->cellXfs);
                $styleIx = (int)$this->styles->cellXfs["count"];
                $this->styles->cellXfs["count"] = (int)$this->styles->cellXfs["count"]+1;
                $c["s"] = $styleIx ; // update cell with style
            }
        }
    } else {
        if ($fillIx!==0){
            $xmlXF = simplexml_load_string("<xf borderId=\"0\" fillId=\"{$fillIx}\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\" applyFill=\"1\"/>");
            $this->insertElementByPosition((int)$this->styles->cellXfs["count"], $xmlXF, $this->styles->cellXfs);
            $styleIx = (int)$this->styles->cellXfs["count"];
            $this->styles->cellXfs["count"] = (int)$this->styles->cellXfs["count"]+1;
            $c["s"] = $styleIx ; // update cell with style
        }
    }
    
    return $c;
    
}

/**
 * This function returns fill color of cell located at $cellAddress. Color is returned as W3C hexadecimal value that starts with hash symbol.
 * If cell is not found it throws an exception.
 *
 * WARNING: in current version this function doesn't take into account alfa channel information stored in first 'two bytes' of OpenXML color information string. It presumes that there's always 'FF' mask in alha channel (no transparency).
 * 
 * @param string $cellAddress Cell address, both A1 and R1C1 address formats are acceptable.
 *
 * @return string Color in W3C format.
 *
 * @category Cell decoration
 */
public function getFillColor($cellAddress){

    // locate cell, if no cell - throw exception
    list( $x, $y, $addrA1, $addrR1C1 ) = self::cellAddress($cellAddress);
    $c = &$this->locateCell($x, $y);
    
    if ($c===null){
        throw new eiseXLSX_Exception('cannot apply fill - no sheet at '.$cellAddress);
    }
    
    // locate style, if no style - add
    if ($c["s"]){
        $cellXf = $this->styles->cellXfs->xf[(int)$c["s"]];
        $fillIx = (int)$cellXf["fillId"];
        $fgColor = $this->styles->fills->fill[$fillIx]->patternFill->fgColor;
        if ($fgColor["rgb"])
            return $fgColor["rgb"];
        else {
            if ($fgColor['theme']){
                return $this->getThemeColor($fgColor['theme']);
            } else if($fgColor["indexed"]){
                return self::colorExcel2W3C(self::$arrIndexedColors[(int)$fgColor["indexed"]]);
            }
        }
        $color = $this->styles->fills->fill[$fillIx]->patternFill->fgColor["rgb"];
        if ($color)
            return $color;
        else 
            return '#FFFFFF';
    } else {
        return '#FFFFFF';
    }
    
    return $c;

}

/**
 * This function returns SRGB color value from theme XML file basing on supplied index value with $theme parameter.
 *
 * @param string $theme Theme index
 *
 * @return hexadecimal SRGB value that correspond given theme index, starting with hash (#) symbol.
 * @ignore
 */
protected function getThemeColor($theme){
    $ixScheme = 0;
    foreach($this->theme->children("a", true)->themeElements[0]->clrScheme[0] as $ix=>$scheme){
        if ((int)$theme==$ixScheme){
            foreach($scheme as $node=>$chl){
                $domch = dom_import_simplexml($chl);
                switch($node) {
                    case "srgbClr":
                    default: 
                        return '#'.$domch->getAttribute("val");
                    case "sysClr":
                        return '#'.$domch->getAttribute("lastClr");
                }
            }
            break;
        }
        $ixScheme++;
    }
}

/**
 * eiseXLSX::cloneRow() method clones row with number $ySrc to $yDest, other existing cells are moved down by one position. 
 * All row contents and styles are simply copied from origin.
 * It returns simpleXML object with cloned row.
 * WARNING: In version 1.6 this method is not friendly to PrintAreas, it they exist and include cells below cloned one. You'll probalby receive 'Currupted file' message from Excel.
 * WARNING: This function wasn't tested for cloning rows from down to up: it's recommended to use it only if $ySrc < $yDest, i.e. when your origin row is upper than destination.
 * 
 * @param int $ySrc - row number of origin. Starts from 1, as user can see on Excel screen
 * @param int $yDest - destination row number.
 *
 * @return object simpleXML object with newly added row
 *
 * @category Sheet manipulations
 */
public function cloneRow($ySrc, $yDest){
    
    $oSrc = $this->locateRow($ySrc);
    if (!$oSrc){
        return null;
    }
    
    $domSrc = dom_import_simplexml($oSrc);
    $oDest = simplexml_import_dom($domSrc->cloneNode(true));
    
    // clean-up <v> and t from cells, update address
    foreach($oDest->c as $c) {
        unset($c["t"]);
        unset($c->v[0]);
        list($x) = self::cellAddress($c["r"]);
        if(preg_match("/^R([0-9]+)C([0-9]+)$/i", $c["r"]))
            $c["r"] = "R{$yDest}C{$x}";
        else 
            $c["r"] = $this->index2letter($x)."{$yDest}";
    }
    
    $oDest["r"] = $yDest;
    
    $retVal =  $this->insertElementByPosition($yDest, $oDest, $this->_cSheet->sheetData);
    
    $this->shiftDownMergedCells($yDest, $ySrc);
    
    return $retVal;
    
}

/**
 * This function returns sheet ID as specified in sheetId attribute of the officeDocument.
 * 
 * @param $name string - sheet name to find
 *
 * @return string - sheet ID if sheet found in current workbook, otherwise false.
 *
 * @category Workbook manipulations
 */
public function findSheetByName($name){
    
    foreach($this->officeDocument->sheets->sheet as $sheet) {
        if ( (string)$sheet['name'] == $name){
            return (string)$sheet["sheetId"];
        }
    }

    return false;

}

/**
 * Function sets sheet with specified $id as active. Internally, $this->_cSheet becomes a sheet with $id.
 * If such sheet cannot be located in the workbook, function throws an exception.
 * NOTICE: sheet numbers (Id's) are started from 1 in speadSheetML.
 * 
 * @param string $id - sheet ID as specified in sheetId attribute of the officeDocument.
 *
 * @return object SimpleXML object that represents the sheet.
 *
 * @category Workbook manipulations
 */
public function selectSheet($id) {
    if(!isset($this->arrSheets[$id])) {
        throw new eiseXLSX_Exception('can\'t select sheet #' . $id);
    }
    $this->_cSheet = $this->arrSheets[$id];
    return $this;
}

/**
 * This method clones original sheet with sheetId supplied with $originSheetId parameter into new one labeled as $newSheetName
 * New sheet doesn't become active. eiseXLSX::cloneSheet() returns sheetId of newly created sheet.
 * 
 * @param string $originSheetId - sheetId of origin sheet
 * @param string $newSheetName - new sheet label, if not set eiseXLSX sets 'Sheet <newSheetId>' as label.
 * 
 * @return string $newSheetId - id of sheet added to the workbook.
 *
 * @category Workbook manipulations
 */
public function cloneSheet($originSheetId, $newSheetName = ''){

    // if origin sheet cannot be located, we throw an exception
    if(!isset($this->arrSheets[$originSheetId])) {
        throw new eiseXLSX_Exception('can\'t select sheet #' . $originSheetId);
    }
    
    // determine maximum sheet ID
    $maxID = 1;$maxSheetFileIX = 1;
    foreach($this->officeDocument->sheets->sheet as $sheet) {
        
        // sheetID
        $maxID = max((int)$sheet["sheetId"], $maxID);
        // filename
        $sheetFileName = basename($this->arrSheetPath[(string)$sheet["sheetId"]]);
        preg_match('/sheet([0-9]+)\.xml/', $sheetFileName, $arrMatch);
        $sheetFileIX = (int)$arrMatch[1];
        $maxSheetFileIX = max($maxSheetFileIX, $sheetFileIX);
        // rel
        $relId = $sheet->attributes('r', TRUE)->id;
        $maxRelID = max($maxRelID, (int)str_replace('rId', '', $relId));
        
    }
    // new sheetID = max +1
    $newSheetID = $maxID+1;
    $newSheetRelID = 'rId'.($maxRelID+1);
    $newSheetFileName = 'sheet'.($maxSheetFileIX+1).'.xml';
    $newSheetFullPath = dirname($this->arrSheetPath[(string)$originSheetId]).'/'.$newSheetFileName;
    $newSheetName = ($newSheetName ? $newSheetName : "Sheet {$newSheetID}");
       
    // clone XLSX by source
    $this->arrXMLs[$newSheetFullPath] = simplexml_load_string($this->arrSheets[(string)$originSheetId]->asXML());
    // clone XLSX by source
    if (isset($this->arrXMLs[self::getRelFilePath($this->arrSheetPath[$originSheetId])]))
        $this->arrXMLs[self::getRelFilePath($newSheetFullPath)] = 
            simplexml_load_string($this->arrXMLs[self::getRelFilePath($this->arrSheetPath[$originSheetId])]->asXML());
    
    // add data to office document
    $newSh = $this->officeDocument->sheets->addChild('sheet');
    $newSh->addAttribute('r:id', $newSheetRelID, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
    $newSh->addAttribute('sheetId', $newSheetID);
    $newSh->addAttribute('name', $newSheetName);
    
    // add data to rels
    // <Relationship Target="worksheets/sheet5.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId5"/>
    $newRel = $this->arrXMLs[$this->officeDocumentRelPath]->addChild('Relationship');
    $newRel->addAttribute('Target', 'worksheets/'.$newSheetFileName);
    $newRel->addAttribute('Type', "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
    $newRel->addAttribute('Id', $newSheetRelID);
    
    // add data to content types
    //<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet5.xml"/>
    $newOvr = $this->arrXMLs["/[Content_Types].xml"]->addChild('Override');
    $newOvr->addAttribute('ContentType', "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    $newOvr->addAttribute('PartName', "/xl/worksheets/".$newSheetFileName );
    
    // recalc worksheet links
    $this->updateWorkbookLinks();

    return (string)$newSheetID;

}

/**
 * This method changes sheet tab label for specified sheet with $sheetId to $newName.
 *
 * @param string $sheetId sheetId of the sheet to be renamed
 * @param string $newName new sheet tab label
 * 
 * @return null
 *
 * @category Workbook manipulations
 */
public function renameSheet($sheetId, $newName){

    // if target sheet cannot be located, we throw an exception
    if(!isset($this->arrSheets[$sheetId])) {
        throw new eiseXLSX_Exception('can\'t get sheet #' . $sheetId);
    }

    foreach($this->officeDocument->sheets->sheet as $sheet) {
        //<sheet r:id="rId1" sheetId="1" name="ACT"/>
        if ((string)$sheet["sheetId"]==(string)$sheetId) {
            $sheet['name'] = $newName;
            break;
        }
    }
    $this->updateAppXML();
}


/**
 * removeSheet() method deletes the sheet specified with $id parameter.
 *
 * @param string $id sheetId of target sheet
 *
 * @return null
 *
 * @category Workbook manipulations
 */
public function removeSheet($id) {
    
    $sheetXMLFileName = $this->arrSheetPath[(string)$id];
    // determine sheet XML rels
    $sheetXMLRelsFileName = self::getRelFilePath($sheetXMLFileName);
    // loop it, delete files
    if ($this->arrXMLs[$sheetXMLRelsFileName]->Relationship)
    foreach($this->arrXMLs[$sheetXMLRelsFileName]->Relationship as $Relationship){
        unset($this->arrXMLs[self::getPathByRelTarget($sheetXMLRelsFileName, $Relationship["Target"])]);
    }
    // unlink sheet rels file
    unset($this->arrXMLs[$sheetXMLRelsFileName]);
    // unlink sheet
    unset($this->arrXMLs[$sheetXMLFileName]);
    unset($this->arrSheets[(string)$id]);
    unset($this->arrSheetPath[(string)$id]);
    
    // rebuild elements from workbook.xml and workbook.xml.rels
    $ix = 0;
    foreach($this->officeDocument->sheets->sheet as $sheet) {
        //<sheet r:id="rId1" sheetId="1" name="ACT"/>
        $relId = $sheet->attributes('r', TRUE)->id; // take old relId
        
        $ixRel = 0;
        foreach($this->arrXMLs[$this->officeDocumentRelPath]->Relationship as $Relationship){
            if ((string)$Relationship["Id"]==$relId)
                break;
            $ixRel++;
        }
        
        if ((string)$sheet["sheetId"]==(string)$id) {
            unset($this->arrXMLs[$this->officeDocumentRelPath]->Relationship[$ixRel]);
            $ixToDel = $ix;
            break;
        }
        
        $ix++;
    }
    unset($this->officeDocument->sheets->sheet[$ixToDel]);
    
    // remove content type ref
    $ixDel = $nCount = 0;
    foreach($this->arrXMLs["/[Content_Types].xml"]->Override as $Override){
        if ((string)$Override["PartName"]==$sheetXMLFileName){
            $ixDel = $nCount;
        }
        $nCount++;
    }
    unset($this->arrXMLs["/[Content_Types].xml"]->Override[$ixDel]);
    
    $this->updateWorkbookLinks();
    
}

/**********************************************/
// XLSX internal file structure manupulation
/**********************************************/
/**
 * @ignore
 */
protected function getPathByRelTarget($relFilePath, $targetPath){
    
    // get directory path of rel file
    $relFileDirectory = preg_replace("/(_rels)$/", "", dirname($relFilePath));
    $arrPath = split("/", rtrim($relFileDirectory, "/"));
    
    // calculate path to target file
    $arrTargetPath = split("/", ltrim($targetPath, "/"));    
    foreach($arrTargetPath as $directory){
        switch($directory){
            case ".":
                break;
            case "..":
                if (isset($arrPath[count($arrPath)-1]))
                    unset($arrPath[count($arrPath)-1]);
                else 
                    throw new Exception("Unable to change directory upwards (..)");
                break;
            default:
                $arrPath[] = $directory;
                break;
                
        }
    }
    
    return implode("/", $arrPath);
}

/** @ignore  */
protected function getRelFilePath($xmlPath){
    return dirname($xmlPath)."/_rels".str_replace(dirname($xmlPath), "", $xmlPath).".rels";
}


/**********************************************/
// sheet data manipulation
/**********************************************/
/** @ignore  */
private function updateSharedString($o_si, $data){
    
    $dom_si = dom_import_simplexml($o_si);
         
    while ($dom_si->hasChildNodes()) {
        $dom_si->removeChild($dom_si->firstChild);
    }
    
    if (!is_object($data)){
        $data = simplexml_load_string("<richText><t>".htmlspecialchars($data)."</t></richText>");
    }
    
    foreach($data->children() as $childNode){
    
        $domInsert = $dom_si->ownerDocument->importNode(dom_import_simplexml($childNode), true);
        $dom_si->appendChild($domInsert);
        
    }
    
    return simplexml_import_dom($o_si);

}

/**
 * eiseXLSX::formatDataRead() function helps to inpreter correctly the numeric value in given cell basing on its $style settings. 
 * In spreadsheetML actual interpretation of number that stores in <c> tag is defined by style attribute. Cell format data then can be obtained from styles.xml document of workbook.
 * Current version of eiseXLSX works correctly with the following data types:
 * - dates: cell data is returned as YYYY-MM-DD string
 * - numbers: cell data is returned as string that actually contains number in latin1 locale.
 * NOTE: Current version works only with just a few format types.
 * 
 * @param string $style - <c style="XX"> attrribute. Should be numeric or empty string.
 * @param string $data - contents of <c> tag.
 *
 * @return string - cell data converted to appropriate format.
 */
/** @ignore  */
private function formatDataRead($style, $data){
    // get style tag
    if ((string)$style=="")
        return (string)$data;
    
    $numFmt = (string)$this->styles->cellXfs->xf[(int)$style]["numFmtId"];
    
    switch ($numFmt){
        case "14": // = 'mm-dd-yy';
        case "15": // = 'd-mmm-yy';
        case "16": // = 'd-mmm';
        case "17": // = 'mmm-yy';
        case "18": // = 'h:mm AM/PM';
        case "19": // = 'h:mm:ss AM/PM';
        case "20": // = 'h:mm';
        case "21": // = 'h:mm:ss';
        case "22": // = 'm/d/yy h:mm';
            return date("Y-m-d", 60*60*24* ($data - self::Date_Bias));
            //return $data
        default: 
            if ((int)$numFmt>=164){ //look for custom format number
                foreach($this->styles->numFmts[0]->numFmt as $o_numFmt){
                    if ((int)$o_numFmt["numFmtId"]==(int)$numFmt){
                        $formatCode = (string)$o_numFmt["formatCode"];
                        if (preg_match("/[dmyh]+/i", $formatCode)){ // CHECK THIS OUT!!! it's just a guess!
                            return date("Y-m-d", 60*60*24* ($data - self::Date_Bias));
                        }
                        break;
                    }
                }
            }
            return $data;
            break;
    }
}

/** @ignore  */
private function addSharedString(&$oCell){

    $ssIndex = count($this->sharedStrings->si);
    
    $oSharedString = $this->sharedStrings->addChild("si", "");
    $this->sharedStrings["uniqueCount"] = $ssIndex+1;
    $this->sharedStrings["count"] = $this->sharedStrings["count"]+1;
    
    $oCell["t"] = "s";
    if (isset($oCell->v[0]))
        $oCell->v[0] = $ssIndex;
    else 
        $oCell->addChild("v", $ssIndex);
    
    return $oSharedString;
}

/**
 * This method converts datetime to Excel format, thanks to Excel::Writer::XLSX::Worksheet.pm (perl).
 * Contributed by gibus [https://github.com/gibus].
 *
 * @param string $date_input String value of data to be set. Should be in YYYY-MM-DD[ hh:mm:ss] format.
 *
 * @return number Value in Excel format, ready to be saved in the cell.
 * @ignore
 */
private function convertDateTime($date_input) {  
    $days    = 0;    # Number of days since epoch
    $seconds = 0;    # Time expressed as fraction of 24h hours in seconds
    $year=$month=$day=0;
    $hour=$min  =$sec=0;
    $date_time = $date_input;
    if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches))
    {
        list($junk,$year,$month,$day) = $matches;
    }
    if (preg_match("/(\d{2}):(\d{2}):(\d{2})/", $date_time, $matches))
    {
        list($junk,$hour,$min,$sec) = $matches;
        $seconds = ( $hour * 60 * 60 + $min * 60 + $sec ) / ( 24 * 60 * 60 );
    }
    
    //using 1900 as epoch, not 1904, ignoring 1904 special case
    # Special cases for Excel.
    if ("$year-$month-$day"=='1899-12-31')  return $seconds      ;    # Excel 1900 epoch
    if ("$year-$month-$day"=='1900-01-00')  return $seconds      ;    # Excel 1900 epoch
    if ("$year-$month-$day"=='1900-02-29')  return 60 + $seconds ;    # Excel false leapday
    
    # We calculate the date by calculating the number of days since the epoch
    # and adjust for the number of leap days. We calculate the number of leap
    # days by normalising the year in relation to the epoch. Thus the year 2000
    # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
    $epoch  = 1900;
    $offset = 0;
    $norm   = 300;
    $range  = $year - $epoch;
    
    # Set month days and check for leap year.
    $leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100)) ) ? 1 : 0;
    $mdays = array( 31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );
    
    # Some boundary checks
    if($year < $epoch || $year > 9999) return $seconds;
    if($month < 1     || $month > 12)  return $seconds;
    if($day < 1       || $day > $mdays[ $month - 1 ]) return $seconds;
    
    # Accumulate the number of days since the epoch.
    $days = $day;    # Add days for current month
    $days += array_sum( array_slice($mdays, 0, $month-1 ) );    # Add days for past months
    $days += $range * 365;                      # Add days for past years
    $days += intval( ( $range ) / 4 );             # Add leapdays
    $days -= intval( ( $range + $offset ) / 100 ); # Subtract 100 year leapdays
    $days += intval( ( $range + $offset + $norm ) / 400 );  # Add 400 year leapdays
    $days -= $leap;                                      # Already counted above
    
    # Adjust for Excel erroneously treating 1900 as a leap year.
    if ($days > 59) { $days++;}
    return $days + $seconds;
}

/**
 * This method formats data for writing and actually writes the data into cell. It doesn't deal with 's' (text) data that stored in sharedStrings.xml.
 *
 * @param string $type Data type. 'd' - datetime string, all other - the number format.
 * @param mixed $data Data to be set. Date/time values should be passed as strings like 'YYYY-MM-DD[ hh:mm:ss]'.
 * @param object $c SimpleXML object that represents <c> element of current sheet*.xml.
 *
 * @return nothing.
 * @ignore
 */
private function formatDataWrite($type, $data, $c){
    
    if (isset($c['t']))
        unset($c['t']);
    
    switch($type){
        case "d":
            $c->v[0] = $this->convertDateTime($data);
            break;
        default:
            $c->v[0] = (string)$data;
            break;
    }
}

/** @ignore  */
private function locateCell($x, $y){
    // locates <c> simpleXMLElement and returns it
    
    $addrA1 = $this->index2letter($x).$y;
    $addrR1C1 = "R{$y}C{$x}";
    
    $row = $this->locateRow($y);
    //*
    if ($row===null) {
        return null;
    };
    //*/
    
    foreach($row->c as $ixC => $c){
        
        if($c["r"]==$addrA1 || $c["r"]==$addrR1C1){
            return $c;
        }
        
    }

    return null;
}

/** @ignore  */
private function addCell($x, $y){
    
    $oValue = null;
    
    $oRow = $this->locateRow($y);
    
    if(!$oRow){
        $oRow = $this->addRow($y, simplexml_load_string("<row r=\"{$y}\"></row>"));
    }
    
    $xmlCell = simplexml_load_string("<c r=\"".$this->index2letter($x).$y."\"></c>");
    $oCell = &$this->insertElementByPosition($x, $xmlCell, $oRow);
    
    return $oCell;
    
}

/** @ignore  */
private function locateRow($y){
    //locates <row> tag with r="$y"
    foreach($this->_cSheet->sheetData->row as $ixRow=>$row){
        if($row["r"]==$y){
            return $row;
        }
    }
    return null;
}

/** @ignore  */
private function addRow($y, $oRow){
    // adds row at position and shifts down all the rows below
    
    $this->shiftDownMergedCells($y);
    
    return $this->insertElementByPosition($y, $oRow, $this->_cSheet->sheetData);
    
}

/** @ignore  */
private function shiftDownMergedCells($yStart, $yOrigin = null){
    
    if (count($this->_cSheet->mergeCells->mergeCell)==0)
        return;
    
    $toAdd = Array();
    
    foreach($this->_cSheet->mergeCells->mergeCell as $mergeCell){
        list($cell1, $cell2) = explode(":", $mergeCell["ref"]);
        
        list($x1, $y1) = self::cellAddress($cell1);
        list($x2, $y2) = self::cellAddress($cell2);
        
        if (max($y1, $y2)>=$yStart && min($y1, $y2)<$yStart){ // if mergeCells are crossing inserted row
            throw new eiseXLSX_Exception("mergeCell {$mergeCell["ref"]} is crossing newly inserted row at {$yStart}");
        }
        
        if (min($y1, $y2)>=$yStart){
            $mergeCell["ref"] = $this->index2letter($x1).($y1+1).":".$this->index2letter($x2).($y2+1);
        }
        
        if ($yOrigin!==null)
            if ($y1==$y2 && $y1==$yOrigin){ // if there're merged cells on cloned row we add new <mergeCell>
                $toAdd[] = $this->index2letter($x1).($yStart).":".$this->index2letter($x2).($yStart);
            }
    }
    
    foreach($toAdd as $newMergeCellRange){
            $newMC = $this->_cSheet->mergeCells->addChild("mergeCell");
            $newMC["ref"] = $newMergeCellRange;
            $this->_cSheet->mergeCells["count"] = $this->_cSheet->mergeCells["count"]+1;
    }
    
}

/** @ignore  */
private function insertElementByPosition($position, $oInsert, $oParent){
    
    $domParent = dom_import_simplexml($oParent);
    $domInsert = $domParent->ownerDocument->importNode(dom_import_simplexml($oInsert), true);
    
    $insertBeforeElement = null;
    $ix = 0;
    
    foreach($domParent->childNodes as $element){
        
        $el_position = $this->getElementPosition($element, $ix) ;
        
        if($position < $el_position){ // if needed element is ahead of current one
            $insertBeforeElement = &$element;
            break;
        }
        
        // else we try to insert element between current and next one
        if ($element->nextSibling!==null && $position <= $this->getElementPosition($element->nextSibling, $ix+1)){
            $insertBeforeElement = &$element->nextSibling;
            break;
        }
        $ix++;
    }
    
    $ix = 0;
    if ($domInsert->nodeName == "row")
        foreach($domParent->childNodes as $element){
            $el_position = $this->getElementPosition($element, $ix) ;
            //shift rows/cells down/right
            if ( $el_position  >= $position ){
                $oElement = simplexml_import_dom($element);
                $oElement["r"] =  $el_position +1; //row 'r' attribute
                foreach($oElement->c as $c){ // cells inside it
                    list($x,$y,$a1,$r1c1) = self::cellAddress($c["r"]);
                    $c["r"] = $c["r"]==$a1 ? self::index2letter($x).($el_position +1) : "R".($el_position +1)."C{$x}";
                }
            }
            $ix++;
        }
    
    
    if ($insertBeforeElement!==null){
        return simplexml_import_dom($domParent->insertBefore($domInsert, $insertBeforeElement));
    } else 
        return simplexml_import_dom($domParent->appendChild($domInsert));
    
}

/** @ignore  */
private function getElementPosition($domXLSXElement, $ix){
    
    if (count($domXLSXElement->attributes)!=0)
        foreach($domXLSXElement->attributes as $ix=>$attr)
            if ($attr->name=="r")
                $strPos = (string)$attr->value;
                    
    switch($domXLSXElement->nodeName){
        case "row":
            return (int)$strPos;
        case "c":
            list($x) = self::cellAddress($strPos);
            return (int)$x;
        default:
            return $ix;
    }
    
}

/** @ignore  */
private function getRow($y){
    $oRow = null;
    foreach($this->_cSheet->sheetData->row as $ixRow=>$row){
        if($row["r"]==$y){
            $oRow = &$row;
            break;
        }
    }
    
    if ($oRow===null){
        $oRow = $this->addRow($y);
    }
    
    return $oRow;
    
}

/**
 * Receives cell address in R1C1 or A1 format and returns address variations as array.
 *
 * Array members are: abscissa, ordinate, A1 and R1C1 -formatted addresses.
 * 
 * @param string $cellAddress both R1C1 and A1 address formats are acceptable. Case-insensitive. Examples: "AI75", "r10c25". 
 *
 * @return array `[ $x, $y, $a1, $r1c1 ]`: 
 * * $x - column number (starting from 1)
 * * $y - row number (starting from 1)
 * * $a1 - cell address in A1 format. "A" in capital case.
 * * $r1c1 - cell address in R1C1. "R" and "C" are capital too.
 *
 * @category Cell address routines
 *
 */
public static function cellAddress($cellAddress){
    
    if(preg_match("/^R([0-9]+)C([0-9]+)$/i", $cellAddress, $arrMatch)){ //R1C1 style
        return Array($arrMatch[2], $arrMatch[1], self::index2letter( $arrMatch[2] ).$arrMatch[1]
        , $cellAddress
        //, "R".self::letter2index(self::index2letter( $arrMatch[1] ))."C$arrMatch[2]"
        );
    } else {
        if (preg_match("/^([a-z]+)([0-9]+)$/i", $cellAddress, $arrMatch)){
            $x = self::letter2index($arrMatch[1]);
            $y = $arrMatch[2];
            return Array($x, $y, $cellAddress, "R{$y}C{$x}");
        }
    }
    
    throw new eiseXLSX_Exception("invalid cell address: {$cellAddress}");
}

/** @ignore */
private static function index2letter($index){
    $nLength = ord("Z")-ord("A")+1;
    $strLetter = "";
    while($index > 0){
        
        $rem = ($index % $nLength==0 ? $nLength : $index % $nLength);
        $strLetter = chr(ord("A")+$rem - 1).$strLetter;
        $index = floor($index/$nLength)-($index % $nLength==0 ? 1 : 0);
        
    }

    return $strLetter;
}

/**
 * @ignore 
 *
 * This static function returns OpenXML color value from HTML's hex value like #RRGGBB supplied with $color parameter. 
 * If color code doesn't match W3C HTML format, it throws an exception. 
 * eiseXLSX::colorExcel2W3C() function provides reverse color conversion.
 *
 * @param string $color Hexadecimal RGB value according to W3C HTML specification that starts with the hashtag, e.g. #AABBCC
 *
 * @return string OpenXML hexadecimal value that can be specified in <fgColor rgb=""> attribute (ARGB). 
 *
 * @example echo self::colorW3CExcel('#00CC99'); // output: 'FF00CC99'
 */ 
private static function colorW3C2Excel($color){
    if (!preg_match('/#[0-9A-F]{2}[0-9A-F]{2}[0-9A-F]{2}/i', $color))
        throw new eiseXLSX_Exception("bad W3C color format: {$color}"); 
    return strtoupper(preg_replace("/^(#)/", "FF", $color));
}

/**
 *
 * @ignore
 *
 * This static function returns W3C color value like #RRGGBB from OpenXML color code supplied with $color parameter.
 * If color code doesn't match OpenXML format, it throws an exception.
 * WARNING: in current version this function doesn't take into account alfa channel information stored in first 'two bytes' of OpenXML color information string. It presumes that there's always 'FF' mask in alha channel (no transparency).
 * eiseXLSX::colorW3CExcel() function provides reverse color conversion.
 *
 * @param string $color Hexadecimal RGB value according to W3C HTML specification that starts with the hashtag, e.g. #AABBCC
 *
 * @return string Hexadecimal value that can be specified in <fgColor rgb=""> attribute (ARGB). 
 *
 * @example echo self::colorExcel2W3C('FF00CC99'); // output: '#00CC99'
 */ 
private static function colorExcel2W3C($color){
    if (!preg_match('/[0-9A-F]{2}[0-9A-F]{2}[0-9A-F]{2}[0-9A-F]{2}/i', $color))
        throw new eiseXLSX_Exception("bad OpenXML color format: {$color}"); 
    return strtoupper(preg_replace("/^([0-9A-F]{2})/i", '#', $color));
}

/** @ignore */
private function letter2index($strLetter){
    $x = 0;
    $nLength = ord("Z")-ord("A")+1;
    for($i = strlen($strLetter)-1; $i>=0;$i--){
    
        $letter = strtoupper($strLetter[$i]);
        $nOffset = ord($letter)-ord("A")+1;
        $x += $nOffset*(pow($nLength, (strlen($strLetter)-1)-$i));
        
    }
    return $x;
}

/** @ignore */
private function updateWorkbookLinks(){
    
    //removing activeTab attribute!
    unset($this->officeDocument->bookViews[0]->workbookView[0]["activeTab"]);
    
    $this->arrSheets = Array();
    $this->arrSheetPath = Array();
    
    // remove all sheets from office document rels XML
    /*
    for($ix = (count($this->arrXMLs[$this->officeDocumentRelPath]->Relationship)-1); $ix>0; $ix--){
        if ((string)$this->arrXMLs[$this->officeDocumentRelPath]->Relationship[$ix]["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"){
            unset($this->arrXMLs[$this->officeDocumentRelPath]->Relationship[$ix]);
        } 
    }
    */
    
    //making sheet index
    $ixSheet = 1;
    foreach($this->officeDocument->sheets->sheet as $sheet){
        //<sheet r:id="rId1" sheetId="1" name="ACT"/>
        
        $oldId = (string)$sheet->attributes('r', TRUE)->id;
        $newId = "rId{$ixSheet}";
        
        $sheet->attributes('r', TRUE)->id = $newId;
        
        foreach($this->arrXMLs[$this->officeDocumentRelPath]->Relationship as $Relationship){
            if ($oldId == (string)$Relationship["Id"] 
                 && (string)$Relationship["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"){
                $Relationship["Id"] = $newId;
                $oldPath = (string)$Relationship["Target"];
                if ($oldId!=$newId) {
                    $newPath = dirname($oldPath)."/sheet{$ixSheet}.xml";
                    $Relationship["Target"] = $newPath; //path in relation
                } else {
                    $newPath = $oldPath;
                }
                break;
            }
        }
        if (!$newPath){
            $newPath = $oldPath = "worksheets/sheet{$ixSheet}.xml";
        }
        
        // rename remainig sheets
        $oldAbsolutePath = self::getPathByRelTarget($this->officeDocumentRelPath, $oldPath);
        $newAbsolutePath = self::getPathByRelTarget($this->officeDocumentRelPath, $newPath);
        if ($oldId!=$newId) 
            $this->renameFile($oldAbsolutePath, $newAbsolutePath);
        
        $this->arrSheets[(string)$sheet["sheetId"]] = &$this->arrXMLs[$newAbsolutePath];
        $this->arrSheetPath[(string)$sheet["sheetId"]] = $newAbsolutePath;
        
        if ($oldId!=$newId){ // rename sheet rels only if sheet is changed
            // rename remaining sheets rels
            $relPath = self::getRelFilePath($oldAbsolutePath);
            if($this->arrXMLs[$relPath]->Relationship)
            foreach($this->arrXMLs[$relPath]->Relationship as $Relationship){
                $oldRelTarget = (string)$Relationship["Target"];
                $newRelTarget = preg_replace("/([0-9]+)\.([a-z0-9]+)/i", $ixSheet.'.\2', $oldRelTarget);
                $Relationship["Target"] = $newRelTarget;
                $this->renameFile(self::getPathByRelTarget($relPath, $oldRelTarget), self::getPathByRelTarget($relPath, $newRelTarget));
            }
            $this->renameFile($relPath, self::getRelFilePath($newAbsolutePath));
        }
        
        $ixSheet++;
    }
    
    // update refs in officeDocumentRelPath
    $ixRel = 0;
    foreach($this->arrXMLs[$this->officeDocumentRelPath]->Relationship as $Relationship){
        if ((string)$Relationship["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
            $this->arrXMLs[$this->officeDocumentRelPath]->Relationship[$ixRel]["Id"] = "rId{$ixSheet}";
        if ((string)$Relationship["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
            $this->arrXMLs[$this->officeDocumentRelPath]->Relationship[$ixRel]["Id"] = "rId".($ixSheet+1);
        if ((string)$Relationship["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
            $this->arrXMLs[$this->officeDocumentRelPath]->Relationship[$ixRel]["Id"] = "rId".($ixSheet+2);
        if ((string)$Relationship["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain")
            $this->arrXMLs[$this->officeDocumentRelPath]->Relationship[$ixRel]["Id"] = "rId".($ixSheet+3);
        if ((string)$Relationship["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink")
            $this->arrXMLs[$this->officeDocumentRelPath]->Relationship[$ixRel]["Id"] = "rId".($ixSheet+4);
        $ixRel++;
    }
    
    $this->updateAppXML();
        
}

/** @ignore */
private function updateAppXML(){

    // update app.xml
    $nSheetsOld = (int)$this->arrXMLs["/docProps/app.xml"]->HeadingPairs->children("vt", true)->vector->variant[1]->i4[0];
    $nAllPartsCount = count($this->arrXMLs["/docProps/app.xml"]->TitlesOfParts->children("vt", true)->vector[0]);
    $nOtherStuffCount = $nAllPartsCount-$nSheetsOld;
    $nSheetsNew = count($this->arrSheets);
    $this->arrXMLs["/docProps/app.xml"]->HeadingPairs->children("vt", true)->vector->variant[1]->i4[0] = $nSheetsNew;
    
    // remove ols sheets
    for ($i=$nSheetsOld-1; $i >= 0; $i--){
        unset($this->arrXMLs["/docProps/app.xml"]->TitlesOfParts->children("vt", true)->vector[0]->lpstr[$i]);
    }
    
    // add new sheets at the beginning of list
    $oParent = $this->arrXMLs["/docProps/app.xml"]->TitlesOfParts->children("vt", true)->vector[0];
    $domParent = dom_import_simplexml($oParent);
    $insertBefore = @dom_import_simplexml($this->arrXMLs["/docProps/app.xml"]->TitlesOfParts->children("vt", true)->vector[0]->lpstr[0]);
    foreach($this->officeDocument->sheets->sheet as $sheet) {
        $xmlLpstr = $oParent->addChild('vt:lpstr', (string)$sheet['name'], 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');
        $domInsert = $domParent->ownerDocument->importNode(dom_import_simplexml($xmlLpstr), true);
        if ($insertBefore!==null)
            $domParent->insertBefore($domInsert, $insertBefore);
        else 
            $domParent->appendChild($domInsert);
    }   
    
    $attr = $this->arrXMLs["/docProps/app.xml"]->TitlesOfParts->children("vt", true)->vector->attributes("", true);
    $attr["size"] = ($nSheetsNew+$nOtherStuffCount);

}

/** @ignore */
private function renameFile($oldName, $newName){
    $this->arrXMLs[$newName] = $this->arrXMLs[$oldName];
    unset($this->arrXMLs[$oldName]);
    
    foreach($this->arrXMLs["/[Content_Types].xml"]->Override as $Override){
        if ((string)$Override["PartName"]==$oldName){
            $Override["PartName"] = $newName;
        }
    }
    
}

/** @ignore */
public function unzipToDirectory($zipFilePath, $targetDirName){

    if (file_exists($targetDirName)){
        self::rmrf($targetDirName);
    }
    
    if(!@mkdir($targetDirName, 0777, true)){

        throw new eiseXLSX_Exception('Unable to create directory to unpack files');

    }


    if(!file_exists($zipFilePath))
        throw new eiseXLSX_Exception("File not found: {$zipFilePath}"); 

    $zip=zip_open($zipFilePath);
    if(!$zip) { 
        throw new eiseXLSX_Exception("Wrong file format: {$zipFilePath}"); 
    }

    while($zip_entry=zip_read($zip)) {
        $strFileName=$targetDirName. self::DS .str_replace("/", self::DS, zip_entry_name($zip_entry));
        $dir = dirname($strFileName);
        if (!file_exists($dir)) mkdir($dir, 0777, true);
        zip_entry_open($zip, $zip_entry);
        $strFile = zip_entry_read($zip_entry, zip_entry_filesize($zip_entry));
        file_put_contents($strFileName, $strFile);
        unset($strFile);
        zip_entry_close($zip_entry);
    }
    zip_close($zip);
    unset($zip);

}

/** @ignore */
private function unzipToMemory($zipFilePath){

    $targetDirName = tempnam(sys_get_temp_dir(), 'eiseXLSX_');
    
    $this->unzipToDirectory($zipFilePath, $targetDirName);

    $eiseXLSX_FS = new eiseXLSX_FS($targetDirName);
    $arrRet = $eiseXLSX_FS->get();
    
    self::rmrf($targetDirName);
    
    return $arrRet;
    
}

/** @ignore deletes directory recursively, like rm -rf */
protected function rmrf($dir){
    
    if(is_dir($dir)){
        $ffs = scandir($dir);
        foreach($ffs as $file) { 
            if ($file == '.' || $file == '..') { continue; } 
            $file = $dir. self::DS .$file;
            if(is_dir($file)) self::rmrf($file); else unlink($file); 
        } 
        rmdir($dir);
    } else {
        unlink($dir);
    }
     
}

/**
 * This method outputs Excel sheet, with the following destination options specified in $dest parameter:
 *
 * 1. D - Excel workbook will be sent to the output as an XLSX file for user download, with "Content-disposition: attachment" header. File name should be specified in $fileName parameter. In case when it empty method will use template file/folder name. Missing ".xlsx" extension will be added.
 * 2. I - Excel workbook is being send out with Content-disposition: inline. It works only with older versions of MSIE/MS Office. It's not recommended to use it. Go for "D" with properly specified filename instead.
 * 3. F - Excel workbook will be saved as file with the name and path specified in $fileName parameter. If there's only file name, it will use current path so remember to chdir() to the location you need. If there's no $fileName, method will save workbook under temporary name. File name will be returned.
 * 4. S (or default) - Method will return workbook file as string. If $fileName parameter is set, workbook will be also saved under this name.
 *
 * There's some smart guess option added for $dest parameter: if you specify only $fileName - omitted $dest will be set to 'D'. If $fileName containes directory separators - omitted $dest will be set to 'F' 
 *
 * Below are the examples of typical usage scenarios:
 * * `$xlsx->Output('my_workbook.xlsx', 'D');` - user will see download prompt with the file named 'my_workbook.xlsx'
 * * `$xlsx->Output('my_workbook.xlsx');` - the same, user will see download prompt with the file named 'my_workbook.xlsx'
 * * `$xlsx->Output('/var/files/my_workbook.xlsx', 'F');` - file will be saved on server
 * * `$xlsx->Output('my_workbook.xlsx', 'F');` - file will be saved at server in current working directory
 * * `$xlsx->Output('/my_workbook.xlsx');` - file will be tried to save on server root
 * * `$my_workbook = $xlsx->Output();` - variable $my_workbook will contain workbook file content. Usable when you need to make mail attachment, for example.
 * 
 * @param string $fileName (optional) File name. If not set, original template name will be used. Missing file extension (.xlsx) will be added automatically.
 * @param string $dest (optional) Destination of method output. See description above.
 *
 * @category Read / Write
 *
 * @return string Workbook file name when $dest="F" or string with workbook content when $dest="S". When $dest="I" or "D" it quits PHP with die(). 
 */
public function Output($fileName = "", $dest = "S") {

    // if filename is set but destination is omitted, the conent will echoed with Content-disposition: attachment
    if($fileName && func_num_args()===1)
        $dest = 'D';

    // if filename contains directory separators, the conent will be attempted to save
    if(preg_match('/['.preg_quote('/\\', '/').']/', $fileName) && func_num_args()===1)
        $dest = 'F';
    
    if(!$fileName || in_array($dest, array("I", "D")) ) {
        $fileNameSrc = $fileName;
        $fileName = tempnam(sys_get_temp_dir(), 'eiseXLSX_');
        $remove = ($dest!=='F');
    }
    
    if(is_writable($fileName) || is_writable(dirname($fileName))) {
       include_once(dirname(__FILE__) . eiseXLSX::DS . 'zipfile.php');
       
       // create archive
       $zip = new zipfile();
       foreach($this->arrXMLs as $xmlFileName => $fileContents) {
            $zip->addFile(
                (is_object($fileContents) ? $fileContents->asXML() : $fileContents)
                , str_replace("/", self::DS, ltrim($xmlFileName, "/"))
                );
                
       }
       file_put_contents($fileName, $zip->file());
       // chmod($this->_eiseXLSXPath, 0777);
    } else {
       throw new eiseXLSX_Exception('could not write to file "' . $fileName . '"');
    }

    switch ($dest){
        case "I":
        case "D":
            if( ini_get('zlib.output_compression') ) { 
                ini_set('zlib.output_compression', 'Off'); 
            }

            // http://ca.php.net/manual/en/function.header.php#76749
            header('Pragma: public'); 
            header("Expires: Sat, 26 Jul 1997 05:00:00 GMT");                  // Date in the past    
            header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); 
            header('Cache-Control: no-store, no-cache, must-revalidate');     // HTTP/1.1 
            header('Cache-Control: pre-check=0, post-check=0, max-age=0');    // HTTP/1.1 
            header("Pragma: no-cache"); 
            header("Expires: 0"); 
            header('Content-Transfer-Encoding: none'); 
//            header('Content-Type: application/vnd.ms-excel;');                 // This should work for IE & Opera 
            header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//            header("Content-type: application/x-msexcel");                    // This should work for the rest 
            if ($dest=="I"){
                header('Content-Disposition: inline"');
            }
            if ($dest=="D"){
                $outFileName = ($fileNameSrc ? basename($fileNameSrc) : $this->defaultFileName);
                if(!preg_match('/\.xlsx$/i', $outFileName))
                    $outFileName .= '.xlsx';
                header("Content-Disposition: attachment; filename*=UTF-8''".rawurlencode($outFileName) );
            }
            readfile($fileName); 
            unlink($fileName);

            die();
        case 'F':
            $r = $fileName;
            break;
        case "S":
        default:
            $r = file_get_contents($fileName);
            break;
    }
        
    if($remove) {
        unlink($fileName);
    }

    return $r;

}

}


/**
 * Throwable class for exceptions.
 * 
 */
class eiseXLSX_Exception extends Exception {
    /**
     * Class constructor, updates message and prints debug backtrace.
     *
     */
    public function __construct($msg) {
          parent::__construct('eiseXLSX error: ' . $msg);
    }
/**
 * Allows to get message directly from the caught exception
 */
    public function __toString() {
        return htmlspecialchars($this->getMessage());
    }
}


/**
 * @ignore 
 * Helper class for XLSX internal file manupulations.
 */
class eiseXLSX_FS {

/** @ignore */
private $path;
/** @ignore */
public $dirs = array();
/** @ignore */
public $filesContent = array();

/** @ignore */
public function __construct($path) {
    $this->path = rtrim($path, eiseXLSX::DS);
    return $this;
}

/** @ignore */
public function get() {
    $this->_scan(eiseXLSX::DS);
    return array($this->dirs, $this->filesContent);
}

/** @ignore */
private function _scan($pathRel) {
    
    if($handle = opendir($this->path . $pathRel)) {
        while(false !== ($item = readdir($handle))) {
            if($item == '..' || $item == '.') {
                continue;
            }
            if(is_dir($this->path . $pathRel . $item)) { 
                $this->dirs[] = ltrim($pathRel, eiseXLSX::DS) . $item;
                $this->_scan($pathRel . $item . eiseXLSX::DS);
            } else {
                $this->filesContent[ltrim($pathRel, eiseXLSX::DS) . $item] = file_get_contents($this->path . $pathRel . $item);
            }
        }
        closedir($handle);
    }
    
}


}
