<?php
/**
*将数据库的数据导出至excel表
*@author 黄增冰
*/
class ArraryToExcel{
    /**
    *@param data    mysql中查询的二维数组数据
    *@param path    PHPExcel的目录路径
    *@param colAttr  列属性
    *@param rowAttr  行属性
    *@param options    属性选项
    *@param excelObj   PHPExcel的对象
    *@param valiData   列的有效性数据
    */
    private $data;
    private $path;
    private $excelObj;
    private $colAttr=array(
            'A'=>array(//列的属性设置
                    'colName'=>'',//第一行的列名
                    'keyName'=>'',//每一列对应的赋值数组的key值
                    'width'=>''   //A列的宽度
                ),
            //可以以 A B C D E F ....递增
           /* 'B'=>array(//列的属性设置
                    'colName'=>'',//第一行的列名
                    'keyName'=>'',//每一列对应的赋值数组的key值
                    'width'=>''   //B列的宽度
                ),
            'C'=>array(//列的属性设置
                    'colName'=>'',//第一行的列名
                    'keyName'=>'',//每一列对应的赋值数组的key值
                    'width'=>''   //C列的宽度
                ),
            'D'=>array(//列的属性设置
                    'colName'=>'',//第一行的列名
                    'keyName'=>'',//每一列对应的赋值数组的key值
                    'width'=>''   //D列的宽
                )
            */
        );
    private $rowAttr=array(
            'firstRowHeight'=>'', //第一行的列名的高度
            'height'=>''         //2-OO无穷行的高度
            );
    private $options=array(
            'excelname'=>'导出excel',  //导出的excel的文件的名称
            'sheettitle'=>'sheet1',    //每个工作薄的标题
            'creater'=>'',             //创建者,
            'lastmodified'=>'',        //最近修改时间
            'title'=>'office xls document',//当前活动的主题
            'subject'=>'office xls document',
            'description'=>'数据导出',
            'keywords'=>'数据导出',
            'category'=>''
        );
    // 有效数据
    private $validData=array();

    /**
    *创建实例自动执行函数
    */
    public function __construct($data,$colAttr,$rowAttr,$options,$validData){
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
        Vendor('PHPExcel');
        $this->excelObj=new \PHPExcel();
        $this->data=$data;
        $this->path=trim($path,'/');
        $this->colAttr=array_merge($this->colAttr,$colAttr);
        $this->rowAttr=array_merge($this->rowAttr,$rowAttr);
        $this->options=array_merge($this->options,$options);
        $this->validData=array_merge($this->validData,$validData);
       
    }
        /**
        *  @param data 从数据库取出来的数组
        *  @param excelname  下载保存的文件名称 
        *  @param sheettitle  脚本的表名称
        *  @param creater   创作者
        *
        */
        //设置要注意顺序，先把各种的格式设置好，最后才存入值
        //设置属性
    public function push(){
        $objPHPExcel=$this->excelObj;
        $objPHPExcel->getProperties()
                    ->setCreator($this->options['creater'])
                    ->setLastModifiedBy($this->options['lastmodified'])
                    ->setTitle($title)
                    ->setSubject($this->options['subject'])
                    ->setDescription($this->options['description'])
                    ->setKeywords($this->options['keywords'])
                    ->setCategory($this->options['category']);

        //设置sheet的name
        $objPHPExcel->getActiveSheet()->setTitle($this->options['sheettitle']);

        
        //设置为excel的第一个表                 
         // $objPHPExcel->setActiveSheetIndex(0);
        //循环设置样色          
            foreach($this->colAttr as $key=>$val){
                //设置每一列的字体居中显示，必须要同时设置水平居中和垂直居中
                $objPHPExcel->getActiveSheet()
                         ->getStyle($key)
                         ->getAlignment()
                         ->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER); 
                $objPHPExcel->getActiveSheet()
                         ->getStyle($key)
                         ->getAlignment()
                         ->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

                //设置第一行列名字段
                $objPHPExcel->getActiveSheet()->setCellValue($key.'1',$val['colName']);

                //设置列宽
                if(isset($val['width'])&&!empty($val['width'])){
                    $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth($val['width']);  
                }else{
                    //自动根据字体的长度确定
                    $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setAutoSize(true);
                }
                
            }
            //设置第一行高度
            if(isset($this->rowAttr['firstRowHeight'])&&!empty($this->rowAttr['firstRowHeight'])){
                        $objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight($this->rowAttr['firstRowHeight']);       
            }

            //循环数组赋值excel单元格
            foreach($this->data as $p=>$v){
                //行数num,第二行开始
                $row=$p+2;

                // 设置数据的有效性
                if(isset($this->validData)&&!empty($this->validData)) {
                    // 总分数据有效性下拉菜单
                    $objValidation1=$objPHPExcel->getActiveSheet()->getCell($this->validData['list1'][0].$row)->getDataValidation();
                    $objValidation1->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST);
                    $objValidation1->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION);
                    $objValidation1->setAllowBlank(false);
                    $objValidation1->setShowInputMessage(true);
                    $objValidation1->setShowErrorMessage(true);
                    $objValidation1->setShowDropDown(true);
                    // $objValidation1->setErrorTitle('Input error');
                    // $objValidation1->setError('Value is not in list.');
                    // $objValidation1->setPromptTitle('Pick from list');
                    // $objValidation1->setPrompt('Please pick a value from the drop-down list.');
                    $objValidation1->setFormula1('"' . $this->validData['list1'][1] . '"');
                    $objPHPExcel->getActiveSheet()->getCell('F'.$row)->setDataValidation($objValidation1); 

                    // 学期数据有效性下拉菜单
                    $objValidation2=$objPHPExcel->getActiveSheet()->getCell($this->validData['list2'][0].$row)->getDataValidation();
                    $objValidation2->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST);
                    $objValidation2->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION);
                    $objValidation2->setAllowBlank(false);
                    $objValidation2->setShowInputMessage(true);
                    $objValidation2->setShowErrorMessage(true);
                    $objValidation2->setShowDropDown(true);
                    // $objValidation2->setErrorTitle('Input error');
                    // $objValidation2->setError('Value is not in list.');
                    // $objValidation2->setPromptTitle('Pick from list');
                    // $objValidation2->setPrompt('Please pick a value from the drop-down list.');
                    $objValidation2->setFormula1('"' . $this->validData['list2'][1] . '"');
                    $objPHPExcel->getActiveSheet()->getCell('G'.$row)->setDataValidation($objValidation2);

                }

                foreach($this->colAttr as $k=>$vo){
                    /**
                    *Excel的第A列，uid是你查出数组的键值，下面以此类推
                    *将数组的值赋值excel的单元格
                    */
                    $objPHPExcel->getActiveSheet()->setCellValue($k.$row, $v[$vo["keyName"]]);
                } 

                /**
                *设置行高
                */    
                if(isset($this->rowAttr['height'])&&!empty($this->rowAttr['height'])){
                    $objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight($this->rowAttr['height']);       
                }                   
            }
                 
        ob_end_clean(); //清除缓冲区,避免乱码
        ob_start(); // Added by me
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$this->options['excelname'].'.xls"');
        header('Cache-Control: max-age=0');
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');
        $objWriter->save('php://output');
        exit;
   }
 }
?>