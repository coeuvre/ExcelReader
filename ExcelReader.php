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
    private $columnMappings;

    private $finished = false;
    private $pos = self::FIRST_ROW;
    private $reader;

    private $columnExisted = array();
    private $headers = null;

    /**
     * @param string    $file           filename of an excel file
     * @param array     $columnDefines  array of {
     *                                      "required"  => true | false,
     *                                      "type"      => "string" | "int" | "float" | "date" | "time"
     *                                      "name"      => string
     *                                      "key"       => string
     *                                  }
     * @param array     $columnMappings array of {
     *                                      "systemName" => string
     *                                      "displayName" => string
     *                                  }
     */
    public function __construct($file, $columnDefines, $columnMappings = null) {
        $this->file = $file;
        $this->columnDefines = $columnDefines;
        $this->columnMappings =  $columnMappings;

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
        return $this->finished;
    }

    /**
     * Read the excel file partially
     *
     * @param integer   $step   optional, how many rows read
     * @return array    {
     *                      'list'  => array of rows
     *                      'warn'  => array of warning messages
     *                      'error' => array of error messages
     *                  }
     */
    public function read($step = 1024) {
        if ($this->finished) {
            return null;
        }

        $result = array(
            'list' => array(),
            'error' => array(),
            'warn' => array(),
        );

        if ($this->pos == self::FIRST_ROW) {
            if ($this->headers == null) {
                $this->readHeaders();
            }
            // skip the first row
            $this->pos += 1;

            // parse the title
            $result['error'] = $this->parseHeaders();

            if (count($result['error'])) {
                $this->finished = true;
                return $result;
            }
        }
        $end = $this->pos + $step;

        // load rows into memory
        $this->reader->setReadFilter(new RowRangedReadFilter($this->pos, $end));
        $excel = $this->reader->load($this->file);

        $worksheet = $excel->getActiveSheet();

        $highestRow = $worksheet->getHighestRow();
        if ($highestRow + 1 < $end) {
            $this->finished = true;
        }

        $highestColumn = $worksheet->getHighestColumn();
        $highestColumn = PHPExcel_Cell::columnIndexFromString($highestColumn);

        // read the data into `result`
        for ($row = $this->pos; $row <= $highestRow; ++$row) {
            $rowData = array();
            $warns = array();
            for ($col = 0; $col <= $highestColumn && $col < count($this->columnDefines) && $this->columnExisted[$col]; ++$col) {
                $cell = $worksheet->getCellByColumnAndRow($col, $row);

                $warn = $this->checkType($cell, $col, $row);
                if ($warn) {
                    $warns[] = $warn;
                }

                $rowData[$this->columnDefines[$col]['key']] = $cell->getValue();
            }
            if (count($warns)) {
                $result['warn'] = array_merge($result['warn'], $warns);
            } else {
                $result['list'][] = $rowData;
            }
        }

        // release the memory
        $excel->disconnectWorksheets();

        $this->pos = $end;
        return $result;
    }

    /**
     * Get the headers of the excel
     *
     * @return array
     */
    public function getHeaders() {
        if ($this->headers == null) {
            assert($this->pos == self::FIRST_ROW);
            $this->readHeaders();
        }
        return $this->headers;
    }

    /**
     * Read the headers into $this->headers
     */
    private function readHeaders() {
        $this->reader->setReadFilter(new RowRangedReadFilter(self::FIRST_ROW, self::FIRST_ROW + 1));
        $excel = $this->reader->load($this->file);

        $worksheet = $excel->getActiveSheet();

        $highestColumn = $worksheet->getHighestColumn();
        $highestColumn = PHPExcel_Cell::columnIndexFromString($highestColumn);

        $this->headers = array();
        for ($col = 0; $col <= $highestColumn; ++$col) {
            $this->headers[] = $worksheet->getCellByColumnAndRow($col, self::FIRST_ROW)->getValue();
        }
    }

    /**
     * Check the title of excel according to `columnDefines`
     *
     * @param $titles array     the titles in excel
     * @return array            error messages
     */
    private function parseHeaders() {
        $errors = array();

        for ($i = 0; $i < count($this->columnDefines); ++$i) {
            if (!isset($this->columnDefines[$i]['key']) || $this->columnDefines[$i]['key'] == '') {
                $this->columnDefines[$i]['key'] = $this->columnDefines[$i]['name'];
            }

            if ($i >= count($this->headers)) {
                $errors[] = "Can't find column `" . $this->columnDefines[$i]['name'] . "`";
                $this->columnExisted[$i] = false;
                continue;
            }

            if ($this->columnDefines[$i]['name'] != $this->headers[$i]) {
                if ($this->columnDefines[$i]['required'] == true) {
                    $errors[] = "Can't find column `" . $this->columnDefines[$i]['name'] . "`";
                }
                $this->columnExisted[$i] = false;
            } else {
                $this->columnExisted[$i] = true;
            }
        }

        return $errors;
    }

    /**
     * Check the data type of a cell according to `columnDefines`
     *
     * @param $cell PHPExcel_Cell   cell to be checked
     * @param $col  integer         col index
     * @param $row  integer         row index
     * @return string               warning message
     */
    private function checkType($cell, $col, $row) {
        $warn = '';
        if ($this->columnDefines[$col]['required']) {
            if (self::isTypeNull($cell)) {
                $warn = "[$row, " . $this->columnDefines[$col]['name'] . "] can't be NULL";
            } else {
                switch ($this->columnDefines[$col]['type']) {
                    case "string":
                        if (!self::isTypeString($cell)) {
                            $warn = "[$row, " . $this->columnDefines[$col]['name'] . "] must be STRING";
                        }
                        break;
                    case "int":
                        if (!self::isTypeInt($cell)) {
                            $warn = "[$row, " . $this->columnDefines[$col]['name'] . "] must be INT";
                        }
                        break;
                    case "float":
                        if (!self::isTypeFloat($cell)) {
                            $warn = "[$row, " . $this->columnDefines[$col]['name'] . "] must be FLOAT";
                        }
                        break;
                    case "date":
                        if (!self::isTypeDate($cell)) {
                            $warn = "[$row, " . $this->columnDefines[$col]['name'] . "] must be DATE";
                        }
                        break;
                    case "time":
                        if (!self::isTypeTime($cell)) {
                            $warn = "[$row, " . $this->columnDefines[$col]['name'] . "] must be TIME";
                        }
                        break;
                }

                if ($warn) {
                    $warn = $warn . " (which is `" . $cell->getValue() . "`)";
                }
            }
        }

        return $warn;
    }

    /**
     * Whether the data type of cell is NULL
     *
     * @param $cell PHPExcel_Cell
     * @return bool
     */
    private static function isTypeNull($cell) {
        $value = strtoupper($cell->getValue());
        return ($cell->getDataType() == PHPExcel_Cell_DataType::TYPE_NULL || $value == "NULL" || $value == "");
    }

    /**
     * Whether the data type of cell is string
     *
     * @param $cell PHPExcel_Cell
     * @return bool
     */
    private static function isTypeString($cell) {
        return !self::isTypeNull($cell);
    }

    /**
     * Whether the data type of cell is integer
     *
     * @param $cell PHPExcel_Cell
     * @return bool
     */
    private static function isTypeInt($cell) {
        if (self::isTypeFloat($cell)) {
            $value = $cell->getValue();
            return !preg_match("/\./", $value);
        } else {
            return false;
        }
    }

    /**
     * Whether the data type fo cell is float (including integer)
     *
     * @param $cell PHPExcel_Cell
     * @return bool
     */
    private static function isTypeFloat($cell) {
        return ($cell->getDataType() == PHPExcel_Cell_DataType::TYPE_NUMERIC || is_numeric($cell->getValue()));
    }

    /**
     * Whether the data type of cell is date
     *
     * Data is a string formatted like "YYYY/MM/DD" with optional time "HH:MM:SS"
     *
     * @param $cell PHPExcel_Cell
     * @return bool
     */
    private static function isTypeDate($cell) {
        $subject = $cell->getValue();
        $pattern = "/^[0-9]{4}(\-|\/)[0-9]{1,2}(\\1)[0-9]{1,2}(|\s+[0-9]{1,2}(|:[0-9]{1,2}(|:[0-9]{1,2})))$/";
        if (preg_match($pattern, $subject) && strtotime($subject)) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Whether the data type of cell is time
     *
     * Time is a string formatted like "YYYY/MM/DD HH:MM:SS"
     *
     * @param $cell PHPExcel_Cell
     * @return bool
     */
    private static function isTypeTime($cell) {
        $subject = $cell->getValue();
        $pattern = "/^[0-9]{4}(\-|\/)[0-9]{1,2}(\\1)[0-9]{1,2}\s+[0-9]{1,2}:[0-9]{1,2}:[0-9]{1,2}$/";
        if (preg_match($pattern, $subject) && strtotime($subject)) {
            return true;
        } else {
            return false;
        }
    }
}

