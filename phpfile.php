<?php
session_start();
header("Content-Type:text/html/charset=utf-8");
$dir=dirname(__FILE__);         //找到当前脚本所在路径
require $dir."/PHPExcel/PHPExcel.php";//引入PHPExcel;

$getClass=$_SESSION['getClass'];
$studyYearTime=$_SESSION['studyYearTime']; //传递学年学期时间
$class_HGL=$_SESSION['class_HGL']; //接收课程合格率
$class_DKJGL=$_SESSION['class_DKJGL']; //接收单科及格率
$stu_num=$_SESSION['stu_num'];  //接收班级人数
$array_XH=$_SESSION['XH'];      //接收学号数组
$array_XM=$_SESSION['XM'];      //接收姓名数组
$array_DYF=$_SESSION['DYF'];      //接收德育分数组
$array_TYF=$_SESSION['TYF'];      //接收体育分数组
$array_ZYF=$_SESSION['ZYF'];      //接收智育分数组
$array_GXF=$_SESSION['GXF'];      //接收个性分数组
$array_ZHF=$_SESSION['ZHF'];      //接收综合分数组
$array_RANK=$_SESSION['RANK'];      //接收排名数组
$array_minScore=$_SESSION['minScore'];      //接收最低单科成绩数组
$array_STF=$_SESSION['STF'];      //接收最低单科成绩数组

//$dir=$_SERVER['DOCUMENT_ROOT']."/phpExcelObjHaveDB/table/";
//$filenameSta="aimclass.xls";
//$filename=$dir.$filenameSta;
//$objPHPExcel=PHPExcel_IOFactory::load($filename);//加载文件
$objPHPExcel=new PHPExcel();            //实例化PHPExcel类,创建一个Excel
$objSheet=$objPHPExcel->getActiveSheet();   //获得当前活动单元格

$objPHPExcel->getActiveSheet()->setCellValue('A1','班级');
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10);//设置单元格长度
$objPHPExcel->getActiveSheet()->setCellValue('B1','学号');
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(10);//设置单元格长度
$objPHPExcel->getActiveSheet()->setCellValue('C1','姓名');
$objPHPExcel->getActiveSheet()->setCellValue('D1','德育分');
$objPHPExcel->getActiveSheet()->setCellValue('E1','智育分');
$objPHPExcel->getActiveSheet()->setCellValue('F1','体育分');
$objPHPExcel->getActiveSheet()->setCellValue('G1','个性发展分');
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(12);//设置单元格长度
$objPHPExcel->getActiveSheet()->setCellValue('H1','综合测评总分');
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(14);//设置单元格长度
$objPHPExcel->getActiveSheet()->setCellValue('I1','班级排名');
$objPHPExcel->getActiveSheet()->setCellValue('J1','最低单科成绩');
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(14);//设置单元格长度
$objPHPExcel->getActiveSheet()->setCellValue('K1','第二课堂学分');
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(14);//设置单元格长度
$objPHPExcel->getActiveSheet()->setCellValue('M1','课程合格率');
$objPHPExcel->getActiveSheet()->setCellValue('M2',$class_HGL);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(14);//设置单元格长度
$objPHPExcel->getActiveSheet()->setCellValue('N1','单科及格率');
$objPHPExcel->getActiveSheet()->setCellValue('N2',$class_DKJGL);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(14);//设置单元格长度
$objPHPExcel->getActiveSheet()->setCellValue('O1',$studyYearTime);

for($i = 0;$i<$stu_num;$i++) {          //在A列填入学号
    $j=$i+2;                            //从A2开始填数据
    $objPHPExcel->getActiveSheet()->setCellValue('A'.$j,$getClass);
}

for($i = 0;$i<$stu_num;$i++) {          //在B列填入学号
    $j=$i+2;                            //从B2开始填数据
    $objPHPExcel->getActiveSheet()->setCellValue('B'.$j,$array_XH[$i]);
}

for($i = 0;$i<$stu_num;$i++) {          //在C列填入姓名
    $j=$i+2;                            //从C2开始填数据
    $objPHPExcel->getActiveSheet()->setCellValue('C'.$j,$array_XM[$i]);
}

for($i = 0;$i<$stu_num;$i++) {          //在D列填入德育分
    $j=$i+2;                            //从D2开始填数据
    $objPHPExcel->getActiveSheet()->setCellValue('D'.$j,$array_DYF[$i]);
}

for($i = 0;$i<$stu_num;$i++) {          //在E列填入智育分
    $j=$i+2;                            //从E2开始填数据
    $objPHPExcel->getActiveSheet()->setCellValue('E'.$j,$array_ZYF[$i]);
}

for($i = 0;$i<$stu_num;$i++) {          //在F列填入体育分
    $j=$i+2;                            //从F2开始填数据
    $objPHPExcel->getActiveSheet()->setCellValue('F'.$j,$array_TYF[$i]);
}

for($i = 0;$i<$stu_num;$i++) {          //在G列填入个性分
    $j=$i+2;                            //从G2开始填数据
    $objPHPExcel->getActiveSheet()->setCellValue('G'.$j,$array_GXF[$i]);
}

for($i = 0;$i<$stu_num;$i++) {          //在H列填入综合分
    $j=$i+2;                            //从H2开始填数据
    $objPHPExcel->getActiveSheet()->setCellValue('H'.$j,$array_ZHF[$i]);
}

for($i = 0;$i<$stu_num;$i++) {          //在I列填入排名
    $j=$i+2;                            //从I2开始填数据
    $objPHPExcel->getActiveSheet()->setCellValue('I'.$j,$array_RANK[$i]);
}

for($i = 0;$i<$stu_num;$i++) {          //在J列填入最低单科成绩
    $j=$i+2;                            //从J2开始填数据
    $objPHPExcel->getActiveSheet()->setCellValue('J'.$j,$array_minScore[$i]);
}

for($i = 0;$i<$stu_num;$i++) {          //在K列填入素拓分
    $j=$i+2;                            //从K2开始填数据
    $objPHPExcel->getActiveSheet()->setCellValue('K'.$j,$array_STF[$i]);
}


$objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');//生成excel文件
browser_export('Excel5','class.xls');                     //输出到浏览器
$objWriter->save('php://output');
session_destroy();


function browser_export($type,$filename){
    header('Content-Type:application/vnd.ms-excel');      //告诉浏览器将要输出excel03文件
    header('Content-Disposition:attachment;filename='.$filename.'');//告诉浏览器将输出文件的名称
    header('Cache-Control:max-age=0');                      //禁止缓存
}

//根据下表获得单元格列位置
function getCells($index){
    $arr=range('A','Z');     //A-Z的数组
    return $arr[$index];
}

?>