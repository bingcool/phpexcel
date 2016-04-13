<?php
/**
*导出成绩模板
*
*/
public function score_export() {
	$classid = I('get.classid');
	$classname = I('get.classname');
	if($classid) {
		$map['identity']=0;
		$map['classid']=intval($classid);
		$data=M('user')->field('usernumber,username')->where($map)->order('id')->select();
		//高属性参数
		foreach ($data as $key => $value) {
			$data[$key]['class'] = $classname;
		}
        $colAttr=array(
            //根据自己到处数据的实际情况添加列名
            
            'A'=>array(//列的属性设置
                    'colName'=>'学生班级',//第一行的列名
                    'keyName'=>'class',//每一列对应的赋值数组的key值
                    'width'=>'30'   //每一列的宽度
                ),
            'B'=>array(//列的属性设置
                    'colName'=>'学生编号',//第一行的列名
                    'keyName'=>'usernumber',//每一列对应的赋值数组的key值
                    'width'=>'20',   //每一列的宽度
                ),

            'C'=>array(//列的属性设置
                    'colName'=>'学生姓名',//第一行的列名
                    'keyName'=>'username',//每一列对应的赋值数组的key值
                    'width'=>'10'   //每一列的宽度
                ),
            'D'=>array(//列的属性设置
                    'colName'=>'考试成绩',//第一行的列名
                    'keyName'=>'score',//每一列对应的赋值数组的key值
                    'width'=>'10'   //每一列的宽度
                ),
            'E'=>array(//列的属性设置
                    'colName'=>'考试名称',//第一行的列名
                    'keyName'=>'testname',//每一列对应的赋值数组的key值
                    'width'=>'30'   //每一列的宽度
                ),
            'F'=>array(//列的属性设置
                    'colName'=>'总分',//第一行的列名
                    'keyName'=>'totalscore',//每一列对应的赋值数组的key值
                    'width'=>'10'   //每一列的宽度
                ),
            'G'=>array(//列的属性设置
                    'colName'=>'学期',//第一行的列名
                    'keyName'=>'term',//每一列对应的赋值数组的key值
                    'width'=>'10'   //每一列的宽度
                )       
        );
        //行属性参数
        $rowAttr=array(
            'firstRowHeight'=>'30', //第一行的列名的高度
            'height'=>'25'         //2-OO无从行的高度
            );
        //excel表的属性参数
        $options=array(
            'excelname'=>$classname.'成绩导入模板',  //导出的excel的文件的名称
            'sheettitle'=>$classname,    //每个工作薄的标题
        );

        // 列对应的有效性数据，即下拉菜单的数据
        $validData=array(
        	'list1'=>array('F','100,120,150'),
        	'list2'=>array('G','第一学期,第二学期')
        	
        );
        Vendor('ArrayToExcel');
        $ArraryToExcel=new \ArraryToExcel($data,$colAttr,$rowAttr,$options,$validData);
        $ArraryToExcel->push();

	}
}

// 导入成绩
public function scoreimport() {
	if($_SERVER['REQUEST_METHOD'] == 'POST'){

		$tid=$this->session->get('userid');

		if($_FILES["file"]["error"] > 0){
		  	echo "Error: " . $_FILES["file"]["error"] . "<br>";
		} 
		//获取文件的名的后缀
		$file_types = explode ( ".", $_FILES["file"]["name"]);
		$file_type = $file_types[count($file_types)-1];

	   //判断文件的格式是否符合格式
		if(strtolower( $file_type ) != "xlsx" && strtolower ( $file_type ) != "xls") {
			// $this->error ( '不是Excel文件，重新上传' );	
			echo json_encode(array("status"=>1,"show"=>'不是Excel文件，重新上传'));	
			return; 
		}

		$tmp_file = $_FILES ['file'] ['tmp_name'];
		$savePath="./Public/upload/";
		//把上传的文件保存到指定的文件夹里,以时间戳命名
		$filename = 'score'.$tid.date('Ymdhis') . "." . $file_type;

		if(!copy( $tmp_file, $savePath . $filename )){
			// $this->error ( '上传失败' );
			echo json_encode(array("status"=>1,"show"=>'上传失败'));
			$content ="导入班级成绩excel文件《".$filename."》失败";
			return;
		}

		//获取导入excel表的内容
		Vendor('ExcelToArray');

		$excelObj=new \ExcelToArray($savePath . $filename);

		$data=$excelObj->read();

		// 测试的信息
		// 测试名称
		$testname=$data[0][4];
		// 总人数
		$count=count($data);
		// 总分
		$totalscore=$data[0][5];
		// 学期
		$term=$data[0][6];
		// 学年班级
		$year_class=$data[0][0];

		foreach($data as $key=>$val) {
			$info[$key]=array(
				'usernumber'=>$val[1],
				'username'=>$val[2],
				'score'=>($val[3] ? $val[3] : '无数据'),
			);
			
		}

		// 要导入的数组的信息
		$scoredata=array(
			'tid'=>$tid,
			'testname'=>$testname,
			'year_class'=>$year_class,
			'term'=>$term,
			'totalscore'=>$totalscore,
			'scoreinfo'=>$info,
			'importData'=>date('Y-m-d H:i:s',time()),
		);

		$offlineScore=new \Think\Model\MongoModel('import_score');

		$insetId=$offlineScore->add($scoredata);
		if($insetId) {
			echo json_encode(array("status"=>1,"show"=>'成绩导入成功'));
		}else {
			echo json_encode(array("status"=>1,"show"=>'成绩导入失败'));
		}
	}
	
}
?>