<?php

namespace Jormin\Excel;

use moonland\phpexcel\Excel as MoonlandExcel;

class Excel extends MoonlandExcel{

    /**
     * @var boolean to exit script after export.
     */
    public $exit = true;

    /**
     * saving the xls file to download or to path
     *
     * @param $sheet
     */
    public function writeFile($sheet)
    {
        if (!isset($this->format))
            $this->format = 'Excel2007';
        $objectwriter = \PHPExcel_IOFactory::createWriter($sheet, $this->format);
        $path = 'php://output';
        if (isset($this->savePath) && $this->savePath != null) {
            $path = $this->savePath . '/' . $this->getFileName();
        }
        $objectwriter->save($path);
        $this->exit && exit();
    }
}