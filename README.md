# EMSExcelTools

专门为快递进行分类筛选的可定制小工具

最近公司作账务问题，很多基础数据人工重复计算筛选太麻烦，每一家的资费又不尽相同，所以就用JAVA编写了一个小工具，后期放在web上，进行在线处理返回一个excel进行下载，便于全公司使用。


主要功能:

1,根据区域和不同重量划分，一区二区三区四区五区和可定制的重量范围。

2,根据不同的标准进行新建工作簿，导出筛选完毕的excel文件。

这次JAVA主要还是用POI啊，感觉还算方便，花了几小时研究了下，还是不错滴!不过筛选功能貌似不兼容2016.........

POI EXCEL的基础知识：

一.POI能够读取的EXCEL文档类型

1,HSSF － 提供读写Microsoft Excel XLS格式档案的功能。

2.XSSF - 提供读写Microsoft Excel OOXML XLSX格式档案的功能。

3,HWPF － 提供读写Microsoft Word DOC格式档案的功能。

4,HSLF － 提供读写Microsoft PowerPoint格式档案的功能。

5,HDGF － 提供读Microsoft Visio格式档案的功能。

6,HPBF － 提供读Microsoft Publisher格式档案的功能。

7,HSMF － 提供读Microsoft Outlook格式档案的功能。

我们平时主要使用HSSF和XSSF这两种，其他基本上很少用到。

二.一分钟快速入门流程。

1.导入POI包.

2.//创建输入流读取文件
  InputStream in = new FileInputStream("文件路径"); 

3.//得到Excel工作簿对象    
  HSSFWorkbook wb = new HSSFWorkbook(in);  

4.//得到Excel工作表对象    
  HSSFSheet sheet = wb.getSheetAt(0);   

5.//得到Excel工作表的行    
  HSSFRow row = sheet.getRow(i);  

6.//得到Excel工作表指定行的单元格    
  HSSFCell cell = row.getCell((short) j);
  
三.特殊需求，如单元格样式和筛选等其他.

1.//创建单元格样式

  HSSFCellStyle my_style = hw.createCellStyle();

2.//设置单元格填充 

  my_style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

3.//设置填充颜色

  my_style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);

4.//设置边框

  my_style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框        
  
  my_style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框   
  
  my_style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框    
  
  my_style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框   
