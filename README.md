# office
excel、word导入导出等组件

## 功能简介
* 基于phpword与phpexcel的组件
* 可处理word文档与excel文档的导入导出

## 安装命令
```composer require jiyull/office```

## 使用demo
```use jiyull\Office;

$xlsName  = "export_blank_list";
$xlsCell  = array(
array('id','序号'),
array('name','填写人姓名'),
array('submit_time','填写时间'),
array('mobile','联系方式'),
array('submit_content','内容查看')
);
$xlsData = [];
$xlsData = [
['id' => 1,
'name' => '叫爸爸',
'submit_time' => "2022年",
"mobile" => "15111112222",
"submit_content" => "内容查看",
]
];
$office = new Office();

$returnData = $office->exportExcel($xlsName,$xlsCell, $xlsData, "http://127.0.0.1");
```

##设置Excel样式：
```
//设置宽度
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(200);   //设置单元格宽度
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);   //内容自适应
//设置align（需要引入PHPExcel_Style_Alignment）
$objPHPExcel->getActiveSheet()->getStyle('A18')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY);//水平方向上两端对齐
$objPHPExcel->getActiveSheet()->getStyle( 'A18')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);    //垂直方向上中间居中
//合并拆分单元格
$objPHPExcel->getActiveSheet()->mergeCells('A28:B28');      // A28:B28合并
$objPHPExcel->getActiveSheet()->unmergeCells('A28:B28');    // A28:B28再拆分
//字体大小、粗体、字体、下划线、字体颜色(需引入PHPExcel_Style_Font、PHPExcel_Style_Color)
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setSize(20);
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setName('Candara');
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setUnderline(PHPExcel_Style_Font::UNDERLINE_SINGLE);
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
//默认字体、大小
$objPHPExcel->getDefaultStyle()->getFont()->setName( 'Arial');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(20);
//背景填充
$objPHPExcel->getActiveSheet()->getStyle( 'A3:E3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$objPHPExcel->getActiveSheet()->getStyle( 'A4:E4')->getFill()->getStartColor()->setARGB('FFC125');
// 单元格密码保护不让修改
$objPHPExcel->getActiveSheet()->getProtection()->setSheet( **true**);  // 为了使任何表保护，需设置为真
$objPHPExcel->getActiveSheet()->protectCells( 'A3:E13', 'PHPExcel' ); // 将A3到E13保护 加密密码是 PHPExcel
$objPHPExcel->getActiveSheet()->getStyle( 'B1')->getProtection()->setLocked(PHPExcel_Style_Protection::PROTECTION_UNPROTECTED); //去掉保护
//给单元格内容设置url超链接
$objActSheet->getCell('E26')->getHyperlink()->setUrl( 'http://www.phpexcel.net');    //超链接url地址
$objActSheet->getCell('E26')->getHyperlink()->setTooltip( 'Navigate to website');  //鼠标移上去连接提示信息
```

##注意事项
在部署上，通常的架构是 nginx + php-fpm，对于Excel中图片比较多的数据导入需要设置加大上传文件的限制和超时时间。后续将会在专栏《面向WEB开发人员的Docker》增加 PHP 运行环境的镜像制作。

在文件上传上，通常会出现 413 request Entity too Large 错误，解决的办法是在 nginx 配置中增加以下配置：

>client_max_body_size  2048m;

相应的 PHP 配置也需要修改，需要修改php.ini ：

>upload_max_filesize = 2048M
post_max_size = 2048M


Excel数据导入，通常会触发504错误，这种情况一般是执行时间太短，涉及的 nginx 配置：

>fastcgi_connect_timeout 600;

php-fpm 中的 www.conf
>request_terminate_timeout = 1800
