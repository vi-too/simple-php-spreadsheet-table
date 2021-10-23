<?php

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Table
{
    /**
     * @var Worksheet
     */
    protected $sheet;

    /**
     * @var array $data
     */
    protected $data;

    /**
     * @var array $headers
     */
    protected $headers;

    /**
     * Column and Row index [Column ID, Row Index] e.g. ['A', 1]
     * @var array
     */
    protected $startingCoord;

    public function __construct(Worksheet $sheet, $data = [], string $pCoordinate = null)
    {
        $this->sheet = $sheet;
        $this->data = $data;
        $this->startingCoord = Coordinate::coordinateFromString($pCoordinate ?? 'A1');
        $this->headers = array_keys($data[0]);
    }

    public function getStartingColumn(): string
    {
        return $this->startingCoord[0];
    }

    public function getStartingRow(): int
    {
        return $this->startingCoord[1];
    }

    public function getLastColumn(): string
    {
        $startIndex = Coordinate::columnIndexFromString($this->getStartingColumn());

        return Coordinate::stringFromColumnIndex((count($this->headers) - 1) + $startIndex);
    }

    public function getLastColumnIndex(): int
    {
        return Coordinate::columnIndexFromString($this->getLastColumn());
    }

    public function getLastRow(): int
    {
        return $this->getStartingRow() + count($this->data);
    }

    public function write()
    {
        $this->writeHeader();
        $this->writeList();
    }

    protected function writeHeader()
    {
        foreach ($this->headers as $key => $value) {
            $this->sheet->setCellValue($this->getCoord($key), $value);
        }
    }

    protected function writeList()
    {
        foreach ($this->data as $row => $data) {
            foreach ($this->headers as $col => $header) {
                if (array_key_exists($header, $data)) {
                    $this->sheet->setCellValue($this->getCoord($col, $row + 1), $data[$header]);
                }
            }
        }
    }

    /**
     * Get the cell coordinate relatie to the starting coord.
     * 
     * @return string
     */
    protected function getCoord($col = 0, $row = 0)
    {
        $col = $col + Coordinate::columnIndexFromString($this->getStartingColumn());
        $row = $row + $this->startingCoord[1];

        return $this->sheet->getCellByColumnAndRow($col, $row)->getCoordinate();
    }
}
