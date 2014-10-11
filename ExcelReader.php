<?php

require_once dirname(__FILE__) . '/Classes/PHPExcel.php';

class RowRangedReadFilter implements PHPExcel_Reader_IReadFilter {
    private $start;
    private $end;

    public function __construct($start, $end) {
        $this->start = $start;
        $this->end = $end;
    }

    /**
     * Should this cell be read?
     *
     * @param integer   $column         String column index
     * @param integer   $row            Row index
     * @param string    $worksheetName  Optional worksheet name
     * @return boolean
     */
    public function readCell($column, $row, $worksheetName = '')
    {
        if ($row >= $this->start && ($this->end <= $this->start || $row < $this->end)) {
            return true;
        }
        return false;
    }
}

class ExcelReader {
    // Note: In PHPExcel column index is 0-based while row index is 1-based. That means 'A1' ~ (0,1)
    const FIRST_ROW = 1;

    private $file;
    private $columnDefines;
    private $step;

    private $isEnd;
    private $pos;
    private $reader;

    /**
     * @param string    $file           filename of an excel file
     * @param array     $columnDefines
     * @param integer   $step           optional, how many rows the reader read each time.
     */
    public function __construct($file, $columnDefines, $step = 1024) {
        $this->file = $file;
        $this->columnDefines = $columnDefines;
        $this->step = $step;

        $this->pos = 0;
        $this->isEnd = false;
        $this->reader = PHPExcel_IOFactory::createReaderForFile($file);
        //$this->reader->setReadDataOnly(true);
    }

    /**
     * Reset the reader, so it can read the excel file again
     */
    public function reset() {
        $this->pos = self::FIRST_ROW;
    }

    /**
     * Whether the reader has read the whole excel file
     *
     * @return bool
     */
    public function finished() {
        return $this->isEnd;
    }

    /**
     * Read the excel file partially
     *
     * @return array
     */
    public function read() {
        $end = $this->pos + $this->step;
        if ($this->pos == self::FIRST_ROW) {
            // the first line is title, read one more row
            $end += 1;
        }

        // load rows into memory
        $this->reader->setReadFilter(new RowRangedReadFilter($this->pos, $end));
        $excel = $this->reader->load($this->file);

        $worksheet = $excel->getActiveSheet();
        $highestRow = $worksheet->getHighestRow();
        if ($highestRow + 1 < $end) {
            $this->isEnd = true;
        }

        $highestColumn = $worksheet->getHighestColumn();
        $highestColumn = PHPExcel_Cell::columnIndexFromString($highestColumn);

        $result = array();

        // read the data into `result`
        //
        // NOTE: Skip the first row which is the title
        for ($row = $this->pos == self::FIRST_ROW ? self::FIRST_ROW + 1 : $this->pos; $row <= $highestRow; ++$row) {
            $rowData = array();
            for ($col = 0; $col <= $highestColumn; ++$col) {
                $rowData[] = $worksheet->getCellByColumnAndRow($col, $row)->getValue();
            }
            $result[] = $rowData;
        }

        // release the memory
        $excel->disconnectWorksheets();

        $this->pos = $end;
        return $result;
    }
}

