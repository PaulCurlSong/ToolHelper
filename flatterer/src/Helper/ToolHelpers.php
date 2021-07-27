<?php 
namespace flatterer\Helper;

use PhpOffice\PhpSpreadsheet\IOFactory;

class ToolHelpers
{
	  /**
    * 读取excel文件内容(该方法解决了合并单元格的问题)
    * @param $file 读取的excel文件 string
    * @param $startPos 开始行 int
    * @param $startCol 开始列 string
    * @param $endCol 结束列 string
    * @return excel数据 array
    */
    public static function xlsxReader($file, $startPos, $startCol, $endCol)
    {
        $reader = IOFactory::load($file);

        // 获取第一个sheet
        $sheet = $reader->getActiveSheet(0);

        $res = [];

        // 存储主单元格数据
        $mainCell = [];

        foreach ($sheet->getRowIterator($startPos) as $row) {

            // 规定从第几行开始读
            $tmp = [];

            foreach ($row->getCellIterator($startCol, $endCol) as $cell) {

                // 遍历每列
                if ($cell->isInMergeRange()) {

                    // 如果是合并单元格
                    if ($cell->isMergeRangeValueCell()) {

                        // 如果是主单元格，则存储(格式 范围-单元格值)
                        $mainCell[$cell->getMergeRange()] = $cell->getFormattedValue();
                        $tmp[] = $cell->getFormattedValue();
                    }

                    else{

                        foreach ($mainCell as $mRange => $mVal) {

                            if ($cell->isInRange($mRange)) {
                                // 在其范围则赋值
                               $tmp[] = $mVal;
                            }
                        }
                    }
                }

                else{
                    $tmp[] = $cell->getFormattedValue();
                }
            }
   
            $res[$row->getRowIndex()] = $tmp;
        }

        return $res;
    }
}