<?php
class ExcelToArray {
    /**
    *@param path string    //vendor下的存放phpexcel文件的路径
    *@param filename   string  //上传后的excel文件的文件名
    *@param ext        string  //文件格式后缀名
    *@param $excelData array    //读取excel表数据存放在数组
    */
    private $path;
    private $filename;
    private $ext;
    private $excelData=array();

    /**
    *实例化执行构造函数
    */
    public function __construct($filename){
        /**
        *导入文件的位置一定要准确，本人是把放在Vendor下
        *把下载的PHPExcel文件夹和PHPExcel.php放在Vendor下.
        */
        /**包含加载PHPExcel.php，完成PHPExcel文件夹的一个Autoloader.php的自动注册
        *if (!defined('PHPEXCEL_ROOT')) {
            define('PHPEXCEL_ROOT', dirname(__FILE__) . '/');
            require(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
        }
        *Autoloader.php的load函数采用以_为符号拆分类名的方式，定义到具体的文件类
        */
        vendor('PHPExcel');
        new \PHPExcel();
        $this->filename=$filename;
        $this->ext=$this->getExt();
        //引入phpexcel类(注意你自己的路径)
        /*Vendor("PHPExcel");  
        Vendor("PHPExcel.IOFactory"); 
        Vendor("PHPExcel.Reader.Excel5"); 
        Vendor("PHPExcel.Reader.Excel2007"); 
        */

    }

    /**
    *获取文件后缀名称
    */
    private function getExt(){
        return end(explode('.', $this->filename));
    }
    /*
    @$filename  文件上传的路径名称，包括到文件格式
    @$file_type  类型
    */
    public function read(){
        if(strtolower($this->ext)=='xls')//判断excel表类型为2003还是2007
            {
                $objReader = \PHPExcel_IOFactory::createReader('Excel5');
            }elseif(strtolower($this->ext)=='xlsx')
            {
                $objReader = \PHPExcel_IOFactory::createReader('Excel2007');
            }
        $objReader->setReadDataOnly(true);
        $objPHPExcel = $objReader->load($this->filename);
        $objWorksheet = $objPHPExcel->getActiveSheet();
        $highestRow = $objWorksheet->getHighestRow();
        $highestColumn = $objWorksheet->getHighestColumn();
        $highestColumnIndex = \PHPExcel_Cell::columnIndexFromString($highestColumn);
        
      //由于excel的第一行是字段，所以我们真正的数据实际第二行开始导入数组的。
        for($row = 2; $row <= $highestRow; $row++) {
            for ($col = 0; $col < $highestColumnIndex; $col++) {
                $this->excelData[$row-2][] = (string)$objWorksheet->getCellByColumnAndRow($col, $row)->getValue();
            }
        }
        //清空缓存
        $objPHPExcel->disconnectWorksheets();
        //删除变量
        unset($objReader, $objPHPExcel, $objWorksheet, $highestColumnIndex);
        return $this->excelData;
    }
}
?>