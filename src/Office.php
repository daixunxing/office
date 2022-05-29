<?php
namespace jiyull;

class Office {
    const EXCEL_TYPE = 1;
    const WORD_TYPE = 2;
    public function __construct($type = self::EXCEL_TYPE){
        switch ($type) {
            case self::EXCEL_TYPE:
                $obj = new \PHPExcel();
                break;
            case self::WORD_TYPE:
                $obj = new \PhpWord();
                break;
            default:
                $obj = new \PHPExcel();
        }
        return $obj;
    }

    /**
     * 导出excel
     * @param $expTitle
     * @param $expCellName
     * @param $expTableData
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     * @throws \PHPExcel_Writer_Exception
     */
    public function exportExcel($expTitle,$expCellName,$expTableData, $serverName){
        //include_once EXTEND_PATH.'PHPExcel/PHPExcel.php';//方法二
        $fileName = $expTitle.date('_YmdHis');//or $xlsTitle 文件名称可根据自己情况设定
        $cellNum = count($expCellName);
        $dataNum = count($expTableData);
        //$objPHPExcel = new PHPExcel();//方法一
        $objPHPExcel = new \PHPExcel();//方法二
        $cellName = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ');
        //$objPHPExcel->getActiveSheet(0)->mergeCells('A1:'.$cellName[$cellNum-1].'1');//合并单元格
        //$objPHPExcel->setActiveSheetIndex(0)->setCellValue('A1', $expTitle.'  Export time:'.date('Y-m-d H:i:s'));
        for($i=0;$i<$cellNum;$i++){
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue($cellName[$i].'1', $expCellName[$i][1]);
        }
        // Miscellaneous glyphs, UTF-8
        for($i=0;$i<$dataNum;$i++){
            for($j=0;$j<$cellNum;$j++){
                if ($cellName[$j] == 'A') {
                    $objPHPExcel->getActiveSheet(0)->setCellValueExplicit($cellName[$j].($i+2), $expTableData[$i][$expCellName[$j][0]], \PHPExcel_Cell_DataType::TYPE_STRING);
                } else {
                    $objPHPExcel->getActiveSheet(0)->setCellValue($cellName[$j].($i+2), $expTableData[$i][$expCellName[$j][0]]);
                }

            }
        }
        ob_end_clean();//这一步非常关键，用来清除缓冲区防止导出的excel乱码
        /*header('pragma:public');
        header('Content-type:application/vnd.ms-excel;charset=utf-8;name="'.$xlsTitle.'.xls"');
        header("Content-Disposition:attachment;filename=$fileName.xls");//"xls"参考下一条备注*/
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');//"Excel2007"生成2007版本的xlsx，"Excel5"生成2003版本的xls
        $dir = 'upload/excel/order/';
        if(!is_dir($dir)) {
            mkdir($dir, 0755,true);
        }
        @$objWriter->save($dir . $fileName.'.xls');//此处加@忽略掉continue 2 的warning
        var_dump($_REQUEST);die;
        return input('server.REQUEST_SCHEME') . '://' . input('server.SERVER_NAME') . '/' . $dir . $fileName.'.xls';
    }


    public function excelToArray ($fileArr) {
        $dataArr = [];
        foreach ($fileArr as $filePath) {
            //加载excel文件
            $objPHPExcelReader = PHPExcel_IOFactory::load($filePath);

            $reader = $objPHPExcelReader->getWorksheetIterator();
            //循环读取sheet
            foreach($reader as $sheet) {
                //读取表内容
                $content = $sheet->getRowIterator();
                //逐行处理
                $resArr = array();
                foreach($content as $key => $items) {

                    $rows = $items->getRowIndex();              //行
                    $columns = $items->getCellIterator();       //列
                    $rowArr = array();
                    //确定从哪一行开始读取
                    if($rows < 2){
                        continue;
                    }
                    //逐列读取
                    foreach($columns as $head => $cell) {
                        //获取cell中数据
                        $data = $cell->getValue();
                        $rowArr[] = $data;
                    }
                    $resArr[] = $rowArr;
                }

            }
            $dataArr = array_merge($dataArr, $resArr) ;
        }
        return $dataArr;

    }

    /*
     * 返回数组中指定多列
     * @param  Array  $input       需要取出数组列的多维数组
     * @param  String $column_keys 要取出的列名，逗号分隔，如不传则返回所有列
     * @param  String $index_key   作为返回数组的索引的列
     * @return Array
    */
    public function arrayColumns($input, $column_keys=null, $index_key=null) {
        $result = array();
        $keys = isset($column_keys) ? explode(',', $column_keys) : array();
        if ($input) {
            foreach ($input as $k => $v) {

                // 指定返回列
                if (!$keys) {
                    $keys = array_keys($v);
                }
                $tmp = array();
                foreach ($keys as $key) {
                    $tmp[$key] = $v[$key];
                }
                // 指定索引列
                if (isset($index_key)) {
                    $result[$v[$index_key]] = $tmp;
                } else {
                    $result[] = $tmp;
                }

            }
        }
        return $result;
    }
}