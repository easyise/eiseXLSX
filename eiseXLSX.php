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
const VERSION = '1.0';
const TPL_DIR = 'templates';
const SKEL_DIR = 'skeleton';
private $_row2col = array();
private $_col2row = array();
private $_cSheet; // current sheet
private $arrSheets = array();
// templates
private $_tpls = array(
    '/[Content_Types].xml' => null,
    '/sheet.xml' => null,
    '/docProps/app.xml' => null,
    '/docProps/core.xml' => null,
    '/xl/_rels/workbook.xml.rels' => null,
    '/xl/workbook.xml' => null,
    '/xl/sharedStrings.xml' => null
);
// parsed templates
private $_parsed = array();
private $arrXMLs = array();

public function __construct( $templatePath='empty' ) {

    $this->_eiseXLSXPath = $path;
    
    // read template
    $templatePath = (file_exists($templatePath) 
        ?  $templatePath 
        :  dirname( __FILE__ ).eiseXLSX::DS. eiseXLSX::TPL_DIR .eiseXLSX::DS.$templatePath
    );
    
    $eiseXLSX_FS = new eiseXLSX_FS($templatePath);
    list($arrDir, $arrFiles) = $eiseXLSX_FS->get();
    $nSheets = 0; $nFirstSheet = 1;
    foreach($arrFiles as $path => $contents) {
        
        $this->arrXMLs[$path] = @simplexml_load_string($contents);
        
        if (!$this->arrXMLs[$path]){
            $this->arrXMLs[$path] = (string)$contents;
        }
        if (preg_match("/sheet([0-9]+)\.xml$/", $path, $arrMatch)){
            $sheetIX = $arrMatch[1];
            $nFirstSheet = ($nSheets==0 ? $sheetIX : $nFirstSheet);
            $this->arrSheets[$sheetIX] = &$this->arrXMLs[$path];
            $nSheets++;
        }
    }
    $this->sharedStrings = &$this->arrXMLs["xl".self::DS."sharedStrings.xml"];
    
    $this->selectSheet($nFirstSheet);
    
}

public function data($cellAddress, $data = null, $t = "s"){
    
    $retVal = null;
    
    list( $x, $y, $addrA1, $addrR1C1 ) = $this->cellAddress($cellAddress);
    
    $c = &$this->locateCell($x, $y);
    if (!$c && $data != null){
        $c = &$this->addCell($x, $y);
    }
    
    if (isset($c->v[0])){ // if it has value
        $o_v = &$c->v[0];
        if ($c["t"]=="s"){ // if existing type is string
            $siIndex = (int)$c->v[0];
            $o_si = &$this->sharedStrings->si[$siIndex];
            $retVal = strip_tags($o_si->asXML()); //return plain string without formatting
        } else { // if not or undefined
            $retVal = (string)$o_v[0];
            if ($data!=null && $t=="s") {// if forthcoming type is string, we add shared string
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
    
    if ($data!=null){ // if we set data
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
                    $c->v[0] = (string)$data;
                    break;
            }
        }
    }
    ///*
    //print_r($this->arrSheets);
    //print_r($oValue);
    //print_r($this->sharedStrings);
    //die();
    //*/
    
    return $retVal;
    
}

private function updateSharedString($o_si, $data){
    
    //echo "<pre>";
    
    $dom_si = dom_import_simplexml($o_si);
    //echo htmlspecialchars($dom_si->C14N())."\r\n";
     
    while ($dom_si->hasChildNodes()) {
        $dom_si->removeChild($dom_si->firstChild);
    }
    
    //echo htmlspecialchars($dom_si->C14N())."\r\n";
    
    
    if (!is_object($data)){
        $data = simplexml_load_string("<richText><t>".htmlspecialchars($data)."</t></richText>");
    }
    
    foreach($data->children() as $childNode){
    
        $domInsert = $dom_si->ownerDocument->importNode(dom_import_simplexml($childNode), true);
        $dom_si->appendChild($domInsert);
        
    }
    
    //echo htmlspecialchars($dom_si->C14N());
    //echo '</pre>';
    return simplexml_import_dom($o_si);
    
    die();
}

private function addSharedString(&$oCell){

    $ssIndex = count($this->sharedStrings->children());
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

private function locateCell($x, $y){
    // locates <c> simpleXMLElement and returns it
    
    $addrA1 = $this->index2letter($x).$y;
    $addrR1C1 = "R{$y}C{$x}";
    
    $row = $this->locateRow($y);
    if ($row==null) {
        echo "{$x} {$y} ". $addrA1." ".$addrR1C1."<br>";
    };
    
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
    
    $oCell = $this->insertElementByPosition($x, $xmlCell, $oRow);
    
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
    
    $retVal =  $this->insertElementByPosition($yDest, $oDest, $this->_cSheet->sheetData)."\r\n";
    
    $this->shiftDownMergedCells($yDest, $ySrc);
    
    return $retVal;
    
}

private function shiftDownMergedCells($yStart, $yOrigin = null){
    
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
        
        if ($yOrigin!=null)
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
    foreach($domParent->childNodes as $element){
        
        $el_position = $this->getElementPosition($element) ;
        
        //shift rows/cells down/right
        if ( $el_position  >= $position ){
            $oElement = simplexml_import_dom($element);
            switch($element->nodeName){
                case "row":
                    $oElement["r"] =  $el_position +1; //row 'r' attribute
                    foreach($oElement->c as $c){ // cells inside it
                        list($x,$y,$a1,$r1c1) = $this->cellAddress($c["r"]);
                        $c["r"] = $c["r"]==$a1 ? $this->index2letter($x).($el_position +1) : "R".($el_position +1)."C{$x}";
                    }
                    break;
                case "c":
                    list($x,$y) = $this->cellAddress($oElement["r"]);
                    $oElement["r"] = $this->index2letter($x+1).$y;
                default: 
                    break;
            }
        }
        
        $nextElement = &$element->nextSibling;
        if (!$nextElement)
            break;
        if ($el_position < $position && $position <= $this->getElementPosition($nextElement)){
            $insertBeforeElement = &$nextElement;
        }
        
        
    }
    
    if ($insertBeforeElement!=null){
        return simplexml_import_dom($domParent->insertBefore($domInsert, $insertBeforeElement));
    } else 
        return simplexml_import_dom($domParent->appendChild($domInsert));
    
}

private function getElementPosition($domXLSXElement){
    
    if ($domXLSXElement->attributes==null)
        throw new eiseXLSX_Exception('no attributes for cell or row');
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
            throw new eiseXLSX_Exception('wrong XLSX tag name"' . $domXLSXElement->nodeName . '"');
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
    
    if ($oRow==null){
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

public function selectSheet($id) {
    if(!isset($this->arrSheets[$id])) {
        throw new eiseXLSX_Exception('can\'t select sheet #' . $id);
    }
    $this->_cSheet = &$this->arrSheets[$id];
    return $this;
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
            //echo $xmlFileName."::".is_object($fileContents)."<br>";
            $zip->addFile(
                (is_object($fileContents) ? $fileContents->asXML() : $fileContents)
                , ltrim($xmlFileName, self::DS)
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

private function _fReplace($tpl, $pattern, $data, $dst = null) {
    if($dst === null) {
        $dst = $tpl;
    }
    $this->_parsed[$dst] = str_replace($pattern, $data, $this->_tpls[$tpl]);
}

private static function _colSort($a, $b) {
    if(($al = strlen($a)) == ($bl = strlen($b))) {
        return strcmp($a, $b);
    }
    return ($al < $bl) ? -1 : 1;
}

private function _parsePairs($val) {
    if(preg_match('#^([a-z]+)(\d+)$#i', $val, $o)) {
        return array($o[1], $o[2]);
    } else {
        return false;
    }
}

}



class eiseXLSX_Exception extends Exception {
    public function __construct($msg) {
          parent::__construct('Simple XLSX error: ' . $msg);
          echo "<pre>";
          debug_print_backtrace();
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