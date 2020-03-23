<?php

/** @noinspection PhpUnhandledExceptionInspection */

use PhpOffice\PhpSpreadsheet\Spreadsheet;

Yii::import('zii.widgets.grid.CGridView');

/**
* @authors Nikola Kostadinov, Dan Straw
* @license MIT License
* @version 1.0
*/
class EExcelView extends CGridView
{
    // Document properties

    /**
     * The creator of the document
     * @var string
     */
    public $creator = '';
    /**
     * Title of the document
     * @var string|null
     */
    public $title = null;
    /**
     * Spreadsheet Subject
     * @var string
     */
    public $subject = '';
    /**
     * Spreadsheet Description
     * @var string
     */
    public $description = '';
    /**
     * Spreadsheet Categories
     * @var string
     */
    public $category = '';
    /**
     * This sheet's title
     * @var string
     */
    public $sheetTitle = 'Sheet1';
    /**
     * Whether the cells are auto width
     * @var bool
     */
    public $autoWidth = true;

    const EXPORT_TYPE_XLS = 'Xls';
    const EXPORT_TYPE_XLSX = 'Xlsx';
    /**
     * The format the file will be created in.
     * See the $writerTypes array for the full list of valid values
     * @var string
     */
    public $exportType = self::EXPORT_TYPE_XLS;

    /**
     * if needed, disable paging to export all data
     * @var bool
     */
    public $disablePaging = true;

    /**
     * The filename to export to
     * @var string|null
     */
    public $filename = null;

    /**
     * Stream output to browser
     * @var bool
     */
    public $stream = true;

    const GRID_MODE_EXPORT = 'export';
    const GRID_MODE_GRID = 'grid';
    /**
     * Whether to display grid ot export it to selected format. Possible values(grid, export)
     * @var string
     */
    public $grid_mode = self::GRID_MODE_GRID;

    /**
     * GET variable name which contains thegrid mode
     * @var string
     */
    public $grid_mode_var = 'grid_mode';

    /**
     * Set sheet as right-to-left
     * @var bool
     */
    public $rtl_sheet = false;

    // buttons config
    public $exportButtonsCSS = 'summary';
    public $exportButtons = array(self::EXPORT_TYPE_XLSX);
    public $exportText = 'Export to: ';

    // callbacks
    public $onRenderHeaderCell = null;
    public $onRenderDataCell = null;
    public $onRenderFooterCell = null;

    // Constants used by the mimeTypes
    const CONTENT_TYPE = 'Content-type';
    const FILE_EXTENSION = 'extension';
    const FILE_CAPTION = 'caption';

    /**
     * Writer Types used for streaming.
     * See PhpOffice\PhpSpreadsheet\IOFactory::writers
     * @var array
     */
    private static $writerTypes = [
        self::EXPORT_TYPE_XLS => array(
            self::CONTENT_TYPE=>'application/vnd.ms-excel',
            self::FILE_EXTENSION=>'xls',
            self::FILE_CAPTION=>'Excel(*.xls)',
        ),
        self::EXPORT_TYPE_XLSX => array(
            self::CONTENT_TYPE=>'application/vnd.ms-excel',
            self::FILE_EXTENSION=>'xlsx',
            self::FILE_CAPTION=>'Excel(*.xlsx)',
        ),
        'Ods' => array(
            self::CONTENT_TYPE => 'application/vnd.oasis.opendocument.spreadsheet',
            self::FILE_EXTENSION=>'ods',
            self::FILE_CAPTION=>'OpenDocument Spreadsheet(*.ods)',
        ),
        'Csv' =>array(
            self::CONTENT_TYPE=>'application/csv',
            self::FILE_EXTENSION=>'csv',
            self::FILE_CAPTION=>'CSV(*.csv)',
        ),
        'Html' => array(
            self::CONTENT_TYPE=>'text/html',
            self::FILE_EXTENSION=>'html',
            self::FILE_CAPTION=>'HTML(*.html)',
        ),
        'Tcpdf' => array(
            self::CONTENT_TYPE=>'application/pdf',
            self::FILE_EXTENSION=>'pdf',
            self::FILE_CAPTION=>'PDF(*.pdf)',
        ),
        'Dompdf' => array(
            self::CONTENT_TYPE=>'application/pdf',
            self::FILE_EXTENSION=>'pdf',
            self::FILE_CAPTION=>'PDF(*.pdf)',
        ),
        'Mpdf' => array(
            self::CONTENT_TYPE=>'application/pdf',
            self::FILE_EXTENSION=>'pdf',
            self::FILE_CAPTION=>'PDF(*.pdf)',
        ),
    ];

    /** @var PhpOffice\PhpSpreadsheet\Spreadsheet $objPhpSpreadsheet */
    protected $objPhpSpreadsheet = null;

    /**
     * @throws CException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function init()
    {
        $this->assignGetParameters();

        if($this->disablePaging) //if needed disable paging to export all data
            $this->dataProvider->pagination = false;

        if($this->grid_mode == self::GRID_MODE_EXPORT &&  ! class_exists("\PhpOffice\PhpSpreadsheet\Spreadsheet") ) {
            $this->grid_mode = self::GRID_MODE_GRID;
            Yii::log('PHPSpreadsheet not found in the vendor directory, please add "phpoffice/phpspreadsheet": ">= 1.10.1" to your composer.json. Export disabled !', CLogger::LEVEL_WARNING, 'EExcelview');
        }

        if($this->grid_mode == self::GRID_MODE_EXPORT)
        {
            $this->title = $this->title ? $this->title : Yii::app()->getController()->getPageTitle();
            $this->initColumns();

            // Create PHPSpreadsheet object
            $this->objPhpSpreadsheet = new Spreadsheet();
            // Setting workbook properties
            $spreadsheetProperties = $this->objPhpSpreadsheet->getProperties();
            $spreadsheetProperties->setCreator($this->creator);
            $spreadsheetProperties->setTitle($this->title);
            $spreadsheetProperties->setSubject($this->subject);
            $spreadsheetProperties->setDescription($this->description);
            $spreadsheetProperties->setCategory($this->category);
            $this->objPhpSpreadsheet->setProperties($spreadsheetProperties);

            if($this->rtl_sheet === true)
                $this->objPhpSpreadsheet->getActiveSheet()->setRightToLeft (true);
        } else {
            // grid mode grid - i.e. render as CGridView.
            parent::init();
        }
    }

    /**
     * Fetches and validates GET parameters passed in
     */
    protected function assignGetParameters()
    {
        if(isset($_GET[$this->grid_mode_var]) && preg_match('/^(export|grid)$/', $_GET[$this->grid_mode_var]) ) {
            $this->grid_mode = $_GET[$this->grid_mode_var];
        }

        if(isset($_GET['exportType']) && key_exists($_GET['exportType'], self::$writerTypes)) {
            $this->exportType = $_GET['exportType'];
        }
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws Exception
     */
    public function renderHeader()
    {
        $a=0;
        foreach($this->columns as $column)
        {
            $a=$a+1;
            if($column instanceof CButtonColumn) {
                $head = $column->header;
            }
            elseif($column->header===null && $column->name!==null)
            {
                if($column->grid->dataProvider instanceof CActiveDataProvider) {
                    $head = $column->grid->dataProvider->model->getAttributeLabel($column->name);
                } else {
                    $head = $column->name;
                }
            } else {
                $head = trim($column->header) !== '' ? $column->header : $column->grid->blankDisplay;
            }

            $cell = $this->objPhpSpreadsheet->getActiveSheet()->setCellValue($this->columnName($a)."1" ,$head);
            if(is_callable($this->onRenderHeaderCell)) {
                call_user_func_array($this->onRenderHeaderCell, array($cell, $head));
            }
        }
    }

    /**
     * @return int
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function renderBody()
    {
        $data=$this->dataProvider->getData();
        $n=count($data);

        if($n>0) {
            for($row=0;$row<$n;++$row) {
                $this->renderRow($row);
            }
        }
        return $n;
    }

    /**
     * @param integer $row
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws Exception
     */
    public function renderRow($row)
    {
        $data=$this->dataProvider->getData();

        $a=0;
        foreach($this->columns as $n=>$column)
        {
            if($column instanceof CLinkColumn)
            {
                if($column->labelExpression!==null)
                    $value=$column->evaluateExpression($column->labelExpression,array('data'=>$data[$row],'row'=>$row));
                else
                    $value=$column->label;
            } elseif($column instanceof CButtonColumn)
                $value = ""; //Dont know what to do with buttons
            elseif($column->value!==null)
                $value=$this->evaluateExpression($column->value ,array('data'=>$data[$row]));
            elseif($column->name!==null) {
                //$value=$data[$row][$column->name];
                $value= CHtml::value($data[$row], $column->name);
                $value=$value===null ? "" : $column->grid->getFormatter()->format($value,'raw');
            }

            $a++;
            $cell = $this->objPhpSpreadsheet->getActiveSheet()->setCellValue($this->columnName($a).($row+2) , strip_tags($value));
            if(is_callable($this->onRenderDataCell))
                call_user_func_array($this->onRenderDataCell, array($cell, $data[$row], $value));
        }
    }

    /**
     * @param $row
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws Exception
     */
    public function renderFooter($row)
    {
        $a=0;
        foreach($this->columns as $n=>$column)
        {
            $a=$a+1;
            if($column->footer)
            {
                $footer =trim($column->footer)!=='' ? $column->footer : $column->grid->blankDisplay;

                $cell = $this->objPhpSpreadsheet->getActiveSheet()->setCellValue($this->columnName($a).($row+2) ,$footer);
                if(is_callable($this->onRenderFooterCell))
                    call_user_func_array($this->onRenderFooterCell, array($cell, $footer));
            }
        }
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws Exception
     */
    public function run()
    {
        if($this->grid_mode == self::GRID_MODE_EXPORT)
        {
            $this->renderHeader();
            $row = $this->renderBody();
            $this->renderFooter($row);

            //set auto width
            if($this->autoWidth) {
                foreach($this->columns as $n=>$column) {
                    $this->objPhpSpreadsheet->getActiveSheet()->getColumnDimension($this->columnName($n + 1))->setAutoSize(true);
                }
            }

            $this->objPhpSpreadsheet->getActiveSheet()->setTitle($this->sheetTitle);

            //create writer for saving
            $writer = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->objPhpSpreadsheet, $this->exportType);
            if(!$this->stream) {
                // save to filename
                $writer->save($this->filename);
            } else {
                //output to browser
                if(!$this->filename)
                    $this->filename = $this->title;
                $this->cleanOutput();
                header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
                header('Pragma: public');
                header('Content-type: '.self::$writerTypes[$this->exportType][self::CONTENT_TYPE]);
                header('Content-Disposition: attachment; filename="'.$this->filename.'.'.self::$writerTypes[$this->exportType][self::FILE_EXTENSION].'"');
                header('Cache-Control: max-age=0');
                $writer->save('php://output');
                ob_start();
                Yii::app()->end();
                ob_end_clean();
            }
        } else
            parent::run();
    }

    /**
     * Returns the coresponding excel column.(Abdul Rehman from yii forum)
     *
     * @param int $index
     * @return string
     * @throws Exception
     */
    public function columnName($index)
    {
        --$index;
        if($index >= 0 && $index < 26)
            return chr(ord('A') + $index);
        else if ($index > 25)
            return ($this->columnName($index / 26)).($this->columnName($index%26 + 1));
            else
                throw new Exception("Invalid Column # ".($index + 1));
    }

    public function renderExportButtons()
    {
        foreach($this->exportButtons as $key=>$button)
        {
            $item = is_array($button) ? CMap::mergeArray(self::$writerTypes[$key], $button) : self::$writerTypes[$button];
            $type = is_array($button) ? $key : $button;
            $url = parse_url(Yii::app()->request->requestUri);
            if (key_exists('query', $url))
                $content[] = CHtml::link($item[self::FILE_CAPTION], '?'.$url['query'].'&exportType='.$type.'&'.$this->grid_mode_var.'=export');
            else
                $content[] = CHtml::link($item[self::FILE_CAPTION], '?exportType='.$type.'&'.$this->grid_mode_var.'=export');
        }
        if($content) {
            echo CHtml::tag('div', array('class' => $this->exportButtonsCSS), $this->exportText . implode(', ', $content));
        }
    }

    /**
    * Performs cleaning on multiple levels.
    *
    * From le_top @ yiiframework.com
    *
    */
    protected static function cleanOutput()
    {
        for($level=ob_get_level();$level>0;--$level)
        {
            @ob_end_clean();
        }
    }

}