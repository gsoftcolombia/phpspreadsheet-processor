<?php

/**
 * Simple php-lib for Importing from or Exporting to MS Excel
 *
 * LIMITATIONS: This lib is able to handle no more than 60 columns in Excel
 *              We think is enought for most of reports but in case you need
 *              more please register an issue in Github. 
 * 
 * @author      Diego Castro <digic93@gmail.com>
 * @author      Omar Ortiz <omaredvega@gmail.com>
 * @author      Oscar Villarraga <dvillarraga@gmail.com>
 * 
 */

use \PhpOffice\PhpSpreadsheet\Style\Conditional;
use \PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use \PhpOffice\PhpSpreadsheet\Style\Alignment;
use \PhpOffice\PhpSpreadsheet\Style\Fill;
use \PhpOffice\PhpSpreadsheet\Style\Border;
use \PhpOffice\PhpSpreadsheet\Shared\Date;
use \PhpOffice\PhpSpreadsheet\Spreadsheet;
use \PhpOffice\PhpSpreadsheet\IOFactory;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx;

Class GsExcelProcessor {

    const CONDITION_CELLIS = Conditional::CONDITION_CELLIS;
    const OPERATOR_GREATERTHAN = Conditional::OPERATOR_GREATERTHAN;
    const OPERATOR_LESSTHAN = Conditional::OPERATOR_LESSTHAN;

    private $sheet;
    private $objPHPExcel;
    private $letters = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ');
    //private $letters array('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26', '27', '28', '29', '30', '31', '32', '33', '34', '35', '36', '37', '38', '39', '40', '41', '42', '43', '44', '45', '46', '47', '48', '49', '50', '51', '52', '52', '53', '54', '55', '56', '57', '58', '59', '60');
    private $iRow;
    private $iColumn;
    private $colorCell;
    private $colorFont;
    private $nameFont;
    private $sizeFont;
    private $bold;
    private $hascenter;
    private $verticalAlignCenter;

    protected $fileName;
    protected $hasBorder;
    protected $conditionals;

    public function __construct()
    {
        $this->iRow = 1;
        $this->iColum = 0;
        $this->colorCell = NULL;
        $this->colorFont = NULL;
        $this->nameFont = NULL;
        $this->sizeFont = NULL;
        $this->hasBorder = true;
        $this->hascenter = false;
        $this->bold = false;
    }
    /* 
     * For Export: First Step
     * @param String $fileName e.g. 'MyFile.xlsx'
    */
    public function setFileName(String $fileName){
        $this->fileName = $fileName;
    }
    /* 
     * For Export: Second Step
     * @param String $creatorName null if the Excel Object is based on a Template.
     * @param String $templatePath There are two possibilities for creating an Excel Object
     *                             1. Empty Excel File, in this case $templatePath must be null
     *                             2. New Excel File based on a Template file, where their path 
     *                                is going to be given in $templatePath
    */
    public function createExcelObject(String $creatorName = null,String $templatePath = null){
        if(is_null($templatePath)){
            $this->objPHPExcel = new Spreadsheet();
            $this->objPHPExcel->getProperties()->setCreator($creatorName)
                ->setLastModifiedBy($creatorName)
                ->setTitle($this->fileName)
                ->setSubject($this->fileName)
                ->setDescription($this->fileName);
        }else{
            try {
                $inputFileType = IOFactory::identify($templatePath);
                $objReader = IOFactory::createReader($inputFileType);
                $this->inputFileType=$inputFileType;
                $objReader->setIncludeCharts(TRUE); //Allows displaying graphics
                $this->objPHPExcel = $objReader->load($templatePath);
            } catch(Exception $e) {
                die('Error loading file "'.pathinfo($templatePath.' || '.$inputFileType,PATHINFO_BASENAME).'": '.$e->getMessage());
            }
        }
    }

    /*
     * For Export: Third Step 
     * Method which push the data into Excel Object
     * @param Array $dataMatrix
     * @param Array $columnTitles (Optional)
    */
    public function setData(array $dataMatrix, array $columnTitles = null) {
        $this->iRow = 1;
        $this->setActiveSheetIndex(0);

        if ($columnTitles != null) {
            $this->setDataByArray($columnTitles, 1);
            $this->iRow ++;
        }

        foreach ($dataMatrix as $value) {
            $this->setDataByArray($value, $this->iRow);
            $this->iRow++;
        }
    }
    
    /*
     * For Export: Fourth Step 
     * Method for export the processed Excel File.
     * @param String $pFilename 'php://output' if you want to have a downloadable file, the path for saving otherwise.
    */
    public function export(String $pFilename){
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $this->fileName . '"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');
        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0
        ob_end_clean();
        $writer =  new Xlsx($this->objPHPExcel);
        $writer->save($pFilename);
        exit;
    }

    /* 
     * For Import: First Step
     * Method for load your Excel File
     * @param String $filePath Excel File location 
    */
    public function loadExcelFile(String $filePath = null){
        try {
            $inputFileType = IOFactory::identify($filePath);
            $objReader = IOFactory::createReader($inputFileType);
            $this->inputFileType=$inputFileType;
            $objReader->setIncludeCharts(TRUE); //Allows displaying graphics
            $this->objPHPExcel = $objReader->load($filePath);
            $this->setActiveSheetIndex(0); // By default it will set the first sheet in the document
        } catch(Exception $e) {
            die('Error loading file "'.pathinfo($filePath.' || '.$inputFileType,PATHINFO_BASENAME).'": '.$e->getMessage());
        }
    }

    /* 
     * For Import: (Optional) Second Step
     * Method for validate your Excel File with custom rules
     * @param Array $rules 'required', 'number', array for in_array validation
     * @return String empty if validation pass, otherwise message
    */

    public function validateRows($rules){
        $rowsCol = $this->getDataFromFile();
        foreach ($rowsCol as $j=>$rows) {
            foreach ($rows as $key =>$value) {
                if(array_key_exists($key, $rules)){
                    if($rules[$key] == 'required'){
                        if(!$value && $value == "")
                            return "Column: ".$this->letters [$key]." Line: ".$rowsCol[$j]['idRow'] ." Field Required";
                    }
                    if($rules[$key] == 'number'){
                        if(!is_numeric($value))
                            return "Column: ".$this->letters [$key]." Line: ".$rowsCol[$j]['idRow'] ." It is not a number";
                    }
                    if(is_array($rules[$key])){
                        if (!in_array($value, $rules[$key]))
                            return "Column: ".$this->letters [$key]." Line: ".$rowsCol[$j]['idRow'] ." No se encontró en el sistema";
                    }
                }
            }
        }        
        return "";
    }

    /*
     * For Import: (Optional) Third Step
     * @param Boolean $includeHeader TRUE if you want to start from row one.
     * @return Array $rows Data Matrix
    */
    public function getDataFromFile(bool $includeHeader = FALSE){
        $highestRow = $this->sheet->getHighestRow();
        $highestColumn = $this->sheet->getHighestColumn();
        $rowBegin=($includeHeader)?1:2;
        $rows=array();
        for ($row = $rowBegin; $row <= $highestRow; $row++){
            $value= $this->sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row,NULL,TRUE,FALSE);
            $value[0]['idRow']=$row;
            $value[0]['is_ok']=1;
            if($value[0][0]||$value[0][1])
                $rows[$row] = $value[0];
        }
        return $rows;
    }

    /*
     * For Import: (Optional) Fourth Step
     * @ToDo Currently depends on Doctrine_Manager Class
     *       It would be nice if SQL Manager Connection is passed as param
     * Method which Call a SQL function/procedure for import your data
     * @param String $procedureName  
     */
    public function loadWithSQLCall($procedureName){
        $rowsCol = $this->getDataFromFile();        
        foreach($rowsCol as &$rowCol){
            foreach($rowCol as $key => $value){
                if(is_null($value))
                    unset($rowCol[$key]);
            }
        }
        foreach ($rowsCol as $key=>$rows) {
            $sql = " CALL $procedureName(";
            $arraInside = array();
            foreach ($rows as $key=>$value) {
                if(is_numeric($key))
                    $arraInside[] = "'". utf8_decode(utf8_encode(str_replace("'", "\'", str_replace(array("\r\n", "\r", "\n", "\0"), NULL, trim($value)))))."'";
            }
            $arraInside = implode(',', $arraInside);
            $sql.=$arraInside.");";
            Doctrine_Manager::getInstance()->getCurrentConnection()->getDbh()->prepare($sql)->execute();                
        }
    }

    public function setActiveSheetIndex($index) {
        $this->objPHPExcel->setActiveSheetIndex($index);
        $this->sheet = $this->objPHPExcel->getActiveSheet();
    }

    public function getActiveSheet() {        
        return $this->objPHPExcel->getActiveSheet();
    }

    public function removeSheetByIndex($index){
        $this->objPHPExcel->removeSheetByIndex($index);
    }

    public function setDataByArray(array $array_values, $iRow, $iColumn = NULL) {
        $this->iColumn = $iColumn == NULL ? 0 : $iColumn;
        foreach ($array_values as $value) {
                $this->setValueCell($value, $iRow, $this->iColumn);
            $this->iColumn ++;
        }
    }

    public function setDataByArrayStyle($array_values, $arrayStyle, $iRow, $iColumn = NULL) {
        $this->iColumn = $iColumn == NULL ? 0 : $iColumn;
        foreach ($array_values as $value) {
                if(isset($arrayStyle[$this->iColumn])){
                    $this->resetStyles();
                    $this->setStyleByArray($arrayStyle[$this->iColumn]);
                }
                $this->setValueCell($value, $iRow, $this->iColumn);
            $this->iColumn ++;
        }
        $this->resetStyles();
    }

    public function setValueCell($value, $iRow, $iColumn,$hasStyle=TRUE) {
        $this->sheet->SetCellValue($this->letters[$iColumn] . $iRow, $value);
        if($hasStyle)
            $this->setStyle($iRow, $iColumn);
    }

    public function getIRow() {
        return $this->iRow;
    }

    public function getIColumn() {
        return $this->iColumn;
    }

    public function setIRow($iRow) {
        $this->iRow = $iRow;
    }

    public function setIColumn($iColumn) {
        $this->iColumn = $iColumn;
    }

    public function setColorCell($colorCell) {
        $this->colorCell = $colorCell;
    }

    public function getVerticalAlignCenter() {
        return $this->verticalAlignCenter;
    }

    public function setVerticalAlignCenter($verticalAlignCenter) {
        $this->verticalAlignCenter = $verticalAlignCenter;
    }

    public function addCountIrow() {
        $this->iRow ++;
        return $this->iRow;
    }

    public function setStyle($iRow, $iColumn) {
        if ($this->colorCell)
            $this->cellColor($iRow, $iColumn);

        if($this->hasBorder)
            $this->cellBorder($iRow, $iColumn);

        if($this->hascenter)
            $this->cellCenter($iRow, $iColumn);

        if($this->verticalAlignCenter)
            $this->cellVerticalAlignCenter($iRow, $iColumn);

        $this->setFont($iRow, $iColumn);
    }

    public function setGeneralFormatNumber(String $pCellCoordinate) {
        $this->objPHPExcel->getActiveSheet()->getStyle($pCellCoordinate)
            ->getNumberFormat()
            ->setFormatCode(NumberFormat::FORMAT_GENERAL);
    }

    protected function setPercentageFormatNumber(String $pCellCoordinate) {
        $this->objPHPExcel->getActiveSheet()->getStyle($pCellCoordinate)
            ->getNumberFormat()
            ->setFormatCode(NumberFormat::FORMAT_PERCENTAGE);
        $this->objPHPExcel->getActiveSheet()->getStyle($pCellCoordinate)
            ->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_RIGHT);
    }

    protected function setDecimalFormatNumber(String $pCellCoordinate) {
        $this->objPHPExcel->getActiveSheet()->getStyle($pCellCoordinate)
            ->getNumberFormat()
            ->setFormatCode(NumberFormat::FORMAT_NUMBER_00);
        $this->objPHPExcel->getActiveSheet()->getStyle($pCellCoordinate)
            ->getAlignment()
            ->setIndent(0)
            ->setHorizontal(Alignment::HORIZONTAL_JUSTIFY);
        $this->objPHPExcel->getActiveSheet()->getStyle($pCellCoordinate)
            ->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_RIGHT);
    }

    public function getLineHighestSheet(){
        $lastRow = $this->sheet->getHighestDataRow();
        $count = 0;
        for($i = $lastRow; $i>=0; $i--){
            $count++;
        }

        return $count+1;
    }

    public function getValue($col, $row){
        return $this->sheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
    }

    public function getColorFont() {
        return $this->colorFont;
    }

    public function getNameFont() {
        return $this->nameFont;
    }

    public function getSizeFont() {
        return $this->sizeFont;
    }

    public function setColorFont($colorFont) {
        $this->colorFont = $colorFont;
    }

    public function setNameFont($nameFont) {
        $this->nameFont = $nameFont;
    }

    public function setSizeFont($sizeFont) {
        $this->sizeFont = $sizeFont;
    }

    public function getBold() {
        return $this->bold;
    }

    public function setBold($bold) {
        $this->bold = $bold;
    }

    public function getCell($col,$row){
        return $this->sheet->getCell($col.$row);
    }

    public function cellColor($iRow, $iColumn, $color = NULL) {
        $color = $color == NULL ? $this->colorCell : $color;

        $this->sheet->getStyle($this->letters[$iColumn] . $iRow)->getFill()->applyFromArray(
            array(
                'type' => Fill::FILL_SOLID,
                'startcolor' => array(
                    'rgb' => $color
                )
            )
        );
    }

    public function setFont($iRow, $iColumn) {
        $font = $this->colorFont != NULL && $this->nameFont != NULL && $this->sizeFont != NULL;

        if ($font) {
            $styleArray = array(
                'font' => array(
                    'bold' => $this->bold,
                    'color' => array('rgb' => $this->colorFont),
                    'size' => $this->sizeFont,
                    'name' => $this->nameFont
            ));

            $this->sheet->getStyle($this->letters[$iColumn] . $iRow)->applyFromArray($styleArray);
        }
    }

    public function mergeCell($iRow, $iColumn, $iRowf, $iColumnf) {
        $this->sheet->mergeCells($this->letters[$iColumn] . $iRow . ':' . $this->letters[$iColumnf] . $iRowf);
    }

    public function cellBorder($iRow, $iColumn, $iRowf = NULL, $iColumnf = NULL) {
        $colum = ($iRowf == NULL && $iColumnf == NULL)? $this->letters[$iColumn] . $iRow : $this->letters[$iColumn] . $iRow. ':' . $this->letters[$iColumnf] . $iRowf ;

        $border_style = array(
            'borders' => array(
                'allborders' => array(
                    'style' => Border::BORDER_THIN,
                    'color' => array('argb' => '000000')
                )
            )
        );

        $this->sheet->getStyle($colum)->applyFromArray($border_style);
    }

    public function cellCenter($iRow, $iColumn, $iRowf = NULL, $iColumnf = NULL) {
        $colum = ($iRowf == NULL && $iColumnf == NULL)? $this->letters[$iColumn] . $iRow : $this->letters[$iColumn] . $iRow. ':' . $this->letters[$iColumnf] . $iRowf ;

        $style = array(
            'alignment' => array(
                'horizontal' => Alignment::HORIZONTAL_CENTER,
            )
        );

        $this->sheet->getStyle($colum)->applyFromArray($style);
    }

    private function createConditional($firstCodition,$secondCondition,$valueCondition,$color){
        $objConditional = new Conditional();
        $objConditional->setConditionType($firstCodition)
                        ->setOperatorType($secondCondition)
                        ->addCondition($valueCondition);
        $objConditional->getStyle()->getFont()->getColor()->setARGB($color);
        return $objConditional;
    }

    public function setConditionalStyle($iRow, $iColumn){
        $conditionalsArray=array();
        foreach($this->conditionals as $conditional){
            $conditionalsArray[] = $this->createConditional($conditional['firstCodition'], $conditional['secondCondition'], $conditional['valueCondition'], $conditional['color']);
        }
        $conditionalStyles = $this->sheet->getStyle($iRow, $iColumn)->getConditionalStyles();
        foreach ($conditionalsArray as $conditional)
            array_push($conditionalStyles, $conditional);
        $this->sheet->getStyle($iRow, $iColumn)->setConditionalStyles($conditionalStyles);
        $this->resetStyles();
    }

    public function cellVerticalAlignCenter($iRow, $iColumn, $iRowf = NULL, $iColumnf = NULL) {
        $colum = ($iRowf == NULL && $iColumnf == NULL)? $this->letters[$iColumn] . $iRow : $this->letters[$iColumn] . $iRow. ':' . $this->letters[$iColumnf] . $iRowf ;

        $style = array(
            'alignment' => array(
                'vertical' => Alignment::VERTICAL_CENTER,
            )
        );

        $this->sheet->getStyle($colum)->applyFromArray($style);
    }

    public function createSheet($index){
        $this->objPHPExcel->createSheet($index);
        $this->sheet = $this->objPHPExcel->setActiveSheetIndex($index);
    }

    public function copyFormatCell($from, $to) {
        $this->sheet->duplicateStyle($this->sheet->getStyle($from), $to);
        $this->sheet->duplicateConditionalStyle($this->sheet->getConditionalStyles($from), $to);
    }

    public function setStyleByArray($style){
        if(isset($style['colorCell']))
            $this->setColorCell($style['colorCell']);
        if(isset($style['colorFont']))
            $this->setColorFont($style['colorFont']);
        if(isset($style['border']))
            $this->setHasBorder($style['border']);
        if(isset($style['center']))
            $this->setHascenter($style['center']);
        if(isset($style['bold']))
            $this->setBold($style['bold']);
        if(isset($style['sizeFont']))
            $this->setSizeFont($style['sizeFont']);
        if(isset($style['verticalAlignCenter']))
            $this->setVerticalAlignCenter($style['verticalAlignCenter']);
        if(isset($style['conditionals'])){
            $this->setConditionals($style['conditionals']);
        }
    }

    public function resetStyles(){
        $this->setColorCell(NULL);
        $this->setColorFont(NULL);
        $this->setHasBorder(FALSE);
        $this->setHascenter(FALSE);
        $this->setHascenter(FALSE);
        $this->setConditionals(FALSE);
    }

    protected function setAutoWidthColumn($letterColumn){
        $this->sheet->getColumnDimension($letterColumn)->setAutoSize(true);
    }

    public function getLetters() {
        return $this->letters;
    }

    public function getConditionals() {
        return $this->conditionals;
    }

    public function setConditionals($conditionals) {
        $this->conditionals = $conditionals;
    }

    public function getHasBorder() {
        return $this->hasBorder;
    }

    public function getHascenter() {
        return $this->hascenter;
    }

    public function setHasBorder($hasBorder) {
        $this->hasBorder = $hasBorder;
    }

    public function setHascenter($hascenter) {
        $this->hascenter = $hascenter;
    }

    public function changeDateToFormat($cellCoordinates,$formatPhpDate){
        $cell = $this->getCell($this->letters[$cellCoordinates['col']],$cellCoordinates['row']);
        $numericDate = $cell->getValue();
        if(is_numeric($numericDate)) {
            $time = Date::excelToTimestamp($numericDate)+86400;
            $date = date($formatPhpDate,$time);
            $this->setValueCell($date, $cellCoordinates['row'], $cellCoordinates['col']);
        }
    }

    public function columnToUppercase($iCol,$limit){
        for($iRow=2;$iRow<=$limit;$iRow++){
            $textCell = $this->getValue($iCol,$iRow);
            $this->setValueCell(strtoupper(trim($textCell)), $iRow, $iCol);
        }
    }
}
