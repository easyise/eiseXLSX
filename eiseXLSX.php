<?php
/****************************************************************/
/*
eiseXLSX class
    
    XLSX file format handling class (Microsoft Office 2007+, spreadsheetML format)
    utilities set:
     - generate filled-in workbook basing on a pre-loaded template
     - save workbook as file
    
    requires SimpleXML
    requires DOM
    
    author: Ilya Eliseev (ie@e-ise.com)
    author: Dmitry Zakharov (dmitry.zakharov@ru.yusen-logistics.com)
    version: 1.0
    
    based on:

     * Simple XLSX [http://www.kirik.ws/eiseXLSX.html]
     * @author kirik [mail@kirik.ws]
     * @version 0.1
     * 
     * Developed under GNU General Public License, version 3:
     * http://www.gnu.org/licenses/lgpl.txt
     
**/
/****************************************************************/
class eiseXLSX {


const DS = DIRECTORY_SEPARATOR;
const Date_Bias = 25569; // number of days between Excel and UNIX epoch
const VERSION = '1.0';
const TPL_DIR = 'templates';

// parsed templates
private $_parsed = array();
private $arrXMLs = array(); // all XML files
private $arrSheets = array();
private $arrSheetPath = array();
private $_cSheet; // current sheet

static $arrIndexedColors = Array('00000000', '00FFFFFF', '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', '0000FFFF', '00000000', '00FFFFFF', '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', '0000FFFF', '00800000', '00008000', '00000080', '00808000', '00800080', '00008080', '00C0C0C0', '00808080', '009999FF', '00993366', '00FFFFCC', '00CCFFFF', '00660066', '00FF8080', '000066CC', '00CCCCFF', '00000080', '00FF00FF', '00FFFF00', '0000FFFF', '00800080', '00800000', '00008080', '000000FF', '0000CCFF', '00CCFFFF', '00CCFFCC', '00FFFF99', '0099CCFF', '00FF99CC', '00CC99FF', '00FFCC99', '003366FF', '0033CCCC', '0099CC00', '00FFCC00', '00FF9900', '00FF6600', '00666699', '00969696', '00003366', '00339966', '00003300', '00333300', '00993300', '00993366', '00333399', '00333333');

public function __construct( $templatePath='empty' ) {

    $this->_eiseXLSXPath = $path;
    
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
        list($arrDir, $arrFiles) = $this->unzip($templatePath);
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


public function data($cellAddress, $data = null, $t = "s"){
    
    $retVal = null;
    
    list( $x, $y, $addrA1, $addrR1C1 ) = $this->cellAddress($cellAddress);
    
    $c = &$this->locateCell($x, $y);
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
        if ($data!=null) // if we'd like to set data
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

public function getRowCount(){
    return count($this->_cSheet->sheetData->row);
}

public function fill($cellAddress, $fillColor){
    // cell address: A1/R1C1
    // fillColor: HTML-style color in Hex pairs, for example: #FFCC66
    
    $fillColor = ($fillColor ? self::colorW3C2Excel($fillColor) : "");
    
    // locate cell, if no cell - throw exception
    list( $x, $y, $addrA1, $addrR1C1 ) = $this->cellAddress($cellAddress);
    $c = &$this->locateCell($x, $y);
    
    if ($c===null){
        throw new eiseXLSX_Exception('cannot apply fill - no sheet at '.$cellAddress);
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

public function getFillColor($cellAddress){

    // locate cell, if no cell - throw exception
    list( $x, $y, $addrA1, $addrR1C1 ) = $this->cellAddress($cellAddress);
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
                return $this->getThemeColor($fgColor['theme'], $fgColor['tint']);
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

function getThemeColor($theme, $tint){
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
            echo htmlspecialchars($scheme[0]->asXML());
            break;
        }
        $ixScheme++;
    }
}

public function cloneRow($ySrc, $yDest){
    // copies row at $ySrc and inserts it at $yDest with shifting down rows below
    
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
        list($x) = $this->cellAddress($c["r"]);
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

public function selectSheet($id) {
    if(!isset($this->arrSheets[$id])) {
        throw new eiseXLSX_Exception('can\'t select sheet #' . $id);
    }
    $this->_cSheet = &$this->arrSheets[$id];
    return $this;
}


public function removeSheet($id) {
    
    $sheetXMLFileName = $this->arrSheetPath[(string)$id];
    // determine sheet XML rels
    $sheetXMLRelsFileName = self::getRelFilePath($sheetXMLFileName);
    // loop it, delete files
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
    
    //remove tab info from app.xml
    $this->arrXMLs["/docProps/app.xml"]->HeadingPairs->children("vt", true)->vector->variant[1]->i4[0] = count($this->arrSheets);
    unset($this->arrXMLs["/docProps/app.xml"]->TitlesOfParts->children("vt", true)->vector->lpstr[$ix]);
    $attr = $this->arrXMLs["/docProps/app.xml"]->TitlesOfParts->children("vt", true)->vector->attributes("", true);
    $attr["size"] = count($this->arrSheets);
    
}

/**********************************************/
// XLSX internal file structure manupulation
/**********************************************/
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

protected function getRelFilePath($xmlPath){
    return dirname($xmlPath)."/_rels".str_replace(dirname($xmlPath), "", $xmlPath).".rels";
}


/**********************************************/
// sheet data manipulation
/**********************************************/
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

//THIS FUNCTION IS IN TODO LIST.
private function formatDataWrite($type, $data, $c){
    
    switch($type){
        case "d":
            $c["s"]=$this->pickupDateStyle();
            $c->v[0] = is_numeric((string)$data) 
                ? (int)$data/60*60*24 + 25569  // if there's number of seconds since UNIX epoch
                : strtotime((string)$data)/60*60*24 + 25569;
            break;
        default:
            $c->v[0] = (string)$data;
            break;
    }
}

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

private function addCell($x, $y){
    
    $oValue = null;
    
    $oRow = $this->locateRow($y);
    
    if(!$oRow){
        $oRow = $this->addRow($y, simplexml_load_string("<row r=\"{$y}\"></row>"));
    }
    
    $xmlCell = simplexml_load_string("<c r=\"".$this->index2letter($x).$y."\" t=\"{$t}\"></c>");
    $oCell = &$this->insertElementByPosition($x, $xmlCell, $oRow);
    
    return $oCell;
    
}

private function locateRow($y){
    //locates <row> tag with r="$y"
    foreach($this->_cSheet->sheetData->row as $ixRow=>$row){
        if($row["r"]==$y){
            return $row;
        }
    }
    return null;
}

private function addRow($y, $oRow){
    // adds row at position and shifts down all the rows below
    
    $this->shiftDownMergedCells($y);
    
    return $this->insertElementByPosition($y, $oRow, $this->_cSheet->sheetData);
    
}

private function shiftDownMergedCells($yStart, $yOrigin = null){
    
    if (count($this->_cSheet->mergeCells->mergeCell)==0)
        return;
    
    $toAdd = Array();
    
    foreach($this->_cSheet->mergeCells->mergeCell as $mergeCell){
        list($cell1, $cell2) = explode(":", $mergeCell["ref"]);
        
        list($x1, $y1) = $this->cellAddress($cell1);
        list($x2, $y2) = $this->cellAddress($cell2);
        
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
                    list($x,$y,$a1,$r1c1) = $this->cellAddress($c["r"]);
                    $c["r"] = $c["r"]==$a1 ? $this->index2letter($x).($el_position +1) : "R".($el_position +1)."C{$x}";
                }
            }
            $ix++;
        }
    
    
    if ($insertBeforeElement!==null){
        return simplexml_import_dom($domParent->insertBefore($domInsert, $insertBeforeElement));
    } else 
        return simplexml_import_dom($domParent->appendChild($domInsert));
    
}

private function getElementPosition($domXLSXElement, $ix){
    
    if (count($domXLSXElement->attributes)!=0)
        foreach($domXLSXElement->attributes as $ix=>$attr)
            if ($attr->name=="r")
                $strPos = (string)$attr->value;
                    
    switch($domXLSXElement->nodeName){
        case "row":
            return (int)$strPos;
        case "c":
            list($x) = $this->cellAddress($strPos);
            return (int)$x;
        default:
            return $ix;
    }
    
}

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

private function cellAddress($cellAddress){
    
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

private function index2letter($index){
    $nLength = ord("Z")-ord("A")+1;
    $strLetter = "";
    while($index > 0){
        
        $rem = ($index % $nLength==0 ? $nLength : $index % $nLength);
        $strLetter = chr(ord("A")+$rem - 1).$strLetter;
        $index = floor($index/$nLength)-($index % $nLength==0 ? 1 : 0);
        
    }

    return $strLetter;
}

//returns XLSX color senternce basing on web's hex like #RRGGBB
private function colorW3C2Excel($color){
    if (!preg_match('/#[0-9A-F]{2}[0-9A-F]{2}[0-9A-F]{2}/i', $color))
        throw new eiseXLSX_Exception("bad W3C color format: {$color}"); 
    return strtoupper(preg_replace("/^(#)/", "FF", $color));
}
//returns W3C hex in #RRGGBB format basing on Excel
private function colorExcel2W3C($color){
    if (!preg_match('/[0-9A-F]{2}[0-9A-F]{2}[0-9A-F]{2}[0-9A-F]{2}/i', $color))
        throw new eiseXLSX_Exception("bad OpenXML color format: {$color}"); 
    return strtoupper(preg_replace("/^([0-9A-F]{2})/i", '#', $color));
}

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


private function updateWorkbookLinks(){
    
    //removing activeTab attribute!
    unset($this->officeDocument->bookViews[0]->workbookView[0]["activeTab"]);
    
    $this->arrSheets = Array();
    $this->arrSheetPath = Array();
    
    //making sheet index
    $ixSheet = 1;
    foreach($this->officeDocument->sheets->sheet as $sheet){
        //<sheet r:id="rId1" sheetId="1" name="ACT"/>
        $oldId = (string)$sheet->attributes('r', TRUE)->id;
        $newId = "rId{$ixSheet}";
        
        $sheet->attributes('r', TRUE)->id = $newId;
        
        foreach($this->arrXMLs[$this->officeDocumentRelPath]->Relationship as $Relationship){
            if ($oldId == (string)$Relationship["Id"]){
                $Relationship["Id"] = $newId;
                $oldPath = (string)$Relationship["Target"];
                if ($oldId!=$newId) {
                    $newPath = dirname($oldPath)."/sheet{$ixSheet}.xml";
                    $Relationship["Target"] = $newPath; //path in relation
                }
                break;
            }
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
        $ixRel++;
    }
    
}

private function renameFile($oldName, $newName){
    $this->arrXMLs[$newName] = $this->arrXMLs[$oldName];
    unset($this->arrXMLs[$oldName]);
    
    foreach($this->arrXMLs["/[Content_Types].xml"]->Override as $Override){
        if ((string)$Override["PartName"]==$oldName){
            $Override["PartName"] = $newName;
        }
    }
    
}

private function unzip($zipFilePath){

    $targetDirName = tempnam(sys_get_temp_dir(), 'eiseXLSX_');
    unlink($targetDirName);
    mkdir($targetDirName, 0777, true);
    
    
    $zip=zip_open($zipFilePath);
    if(!$zip) { 
        throw new eiseXLSX_Exception("Wrong file format"); 
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
    
    $eiseXLSX_FS = new eiseXLSX_FS($targetDirName);
    $arrRet = $eiseXLSX_FS->get();
    
    self::rmrf($targetDirName);
    
    return $arrRet;
    
}

// deletes directory recursively, like rm -rf
protected function rmrf($dir){
    
    $ffs = scandir($dir);
    foreach($ffs as $file) { 
        if ($file == '.' || $file == '..') { continue; } 
        $file = $dir. self::DS .$file;
        if(is_dir($file)) self::rmrf($file); else unlink($file); 
    } 
    rmdir($dir); 
}

public function Output($fileName = "", $dest = "I") {
    
    if(!$fileName) {
       $fileName = tempnam(sys_get_temp_dir(), 'eiseXLSX_');
       $remove = true;
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
                header('Content-Disposition: attachment; filename="' . basename($fileName) . '.xlsx"');
            }
            readfile($fileName); 
            die();
        case "S":
            return file_get_contents($fileName);
        case "F": 
        default:
            break;
    }
        
    if(isset($remove)) {
       unlink($fileName);
    }
}

}



class eiseXLSX_Exception extends Exception {
    public function __construct($msg) {
          parent::__construct('Simple XLSX error: ' . $msg);
          echo "<pre>";
          debug_print_backtrace();
          echo "</pre>";
    }

    public function __toString() {
        return htmlspecialchars($this->getMessage());
    }
}



class eiseXLSX_FS {

private $path;
public $dirs = array();
public $filesContent = array();

public function __construct($path) {
    $this->path = rtrim($path, eiseXLSX::DS);
    return $this;
}

public function get() {
    $this->_scan(eiseXLSX::DS);
    return array($this->dirs, $this->filesContent);
}

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