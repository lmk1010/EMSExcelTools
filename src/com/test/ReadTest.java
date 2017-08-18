package com.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;


public class ReadTest {

	String[] area1 = {"安徽省","浙江省","江苏省","上海市"};
	String[] area2 = {"湖北省","河南省","山东省","江西省","天津市","河北省","湖南省","福建省","山西省","陕西省"};
	String[] area3 = {"北京市","辽宁省","吉林省","重庆市","四川省","广东省","内蒙古自治区","甘肃省","广西省","宁夏省","贵州省","海南省"};
	String[] area4 = {"黑龙江省","云南省"};
	String[] area5 = {"西藏自治区","新疆维吾尔自治区","青海省"};
	
	ArrayList<String[]> Area = new ArrayList<String[]>();
	
	int[] weight = {500,1000,2000,3000,4000,5000};
	
	//工作簿
    String[] sheetname1 = {"一区0-500","一区500-1000","一区1000-2000","一区2000-3000","一区3000-4000","一区4000-5000"};
    String[] sheetname2 = {"二区0-500","二区500-1000","二区1000-2000","二区2000-3000","二区3000-4000","二区4000-5000"};
    String[] sheetname3 = {"三区0-500","三区500-1000","三区1000-2000","三区2000-3000","三区3000-4000","三区4000-5000"};
    String[] sheetname4 = {"四区0-500","四区500-1000","四区1000-2000","四区2000-3000","四区3000-4000","四区4000-5000"};
    String[] sheetname5 = {"五区0-500","五区500-1000","五区1000-2000","五区2000-3000","五区3000-4000","五区4000-5000"};
    
    ArrayList<String[]> sheetname = new ArrayList<String[]>();
	
	public void testarea(){
		Area.add(area1);
		Area.add(area2);
		Area.add(area3);
		Area.add(area4);
		Area.add(area5);
		for (String[] areatest : Area){
		ArrayList<ArrayList<Object>> selectiondata = AreaSelection(areatest,0,500,"g://222.xls");
		for (ArrayList<Object> a:selectiondata){
			System.out.println(a);
		}
		}
	}
	
	public void test9(){
		
	}
	
	@Test
	public void test2() throws IOException{
		
		Area.add(area1);
		Area.add(area2);
		Area.add(area3);
		Area.add(area4);
		Area.add(area5);

		sheetname.add(sheetname1);
		sheetname.add(sheetname2);
		sheetname.add(sheetname3);
		sheetname.add(sheetname4);
		sheetname.add(sheetname5);

		// 0-5 6-11 12-17 18-23 24-
		int temp = 0;
		for (String[] areatest : Area) {
			int mintest = 0;
			int maxtest = 500;
			int count = 0;

			if (maxtest <= 5000) {

				ArrayList<ArrayList<Object>> selectiondata = new ArrayList<ArrayList<Object>>();

				for (String name : sheetname.get(temp)) {
					selectiondata = AreaSelection(areatest, mintest, maxtest,
							"g://excel1.xls");
					WriteNewsheet(selectiondata, name, "g://ceshi3.xls");

					count++;

					if (count == 1) {
						mintest = mintest + 500;
						maxtest = maxtest + 500;
					} else if (count == 2) {
						mintest = mintest + 500;
						maxtest = maxtest + 1000;
					} else if (count > 2) {
						mintest = mintest + 1000;
						maxtest = maxtest + 1000;
					}
				}
				temp++;

			}
			
					
		}
		
		
		
	
	}

	//区域和重量筛选器
	public ArrayList<ArrayList<Object>> AreaSelection(String[] area,double minweight,double maxweight,String selectfilepath){
		
		ArrayList<ArrayList<Object>> selectiondata = new ArrayList<ArrayList<Object>>();
		
		InputStream in;
		try {
			in = new FileInputStream(selectfilepath);
			
			HSSFWorkbook hw = new HSSFWorkbook(in);
			
			HSSFSheet sheet = hw.getSheetAt(0);

			for (int rowindex = 5; rowindex <= sheet.getLastRowNum(); rowindex++) {
				int realrow = rowindex + 1;

				HSSFRow row = sheet.getRow(rowindex);

				HSSFCell areacell = row.getCell(3);
				HSSFCell weightcell = row.getCell(6);

				for (String a : area) {
					if (areacell.getRichStringCellValue().getString().equals(a)) {
						double x = weightcell.getNumericCellValue();
						if ((x > minweight) && (x < maxweight)) {
							ArrayList<Object> list = new ArrayList<Object>(); 
						    list = readbyrow(realrow,selectfilepath);
							selectiondata.add(list);
							int count = 0;
							for (Object b : list) {
								//System.out.print(b + "--");
								count++;
								if (count % 10 == 0) {
									//System.out.println();
								}
							}

						}
					}
				}

			}
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return selectiondata;
		
		
	}
	
	
	//写入新sheet
		public void WriteNewsheet(ArrayList<ArrayList<Object>> selectiondata,String Areasheet,String fileoutpath) throws IOException{
			
			HSSFWorkbook hw;
			File file = new File(fileoutpath);
			
			if (!file.exists()){
				//如果文件不存在则新建
				hw = new HSSFWorkbook();
			}else{
				//如果已存在则直接写入
				InputStream in = new FileInputStream(fileoutpath);
				hw = new HSSFWorkbook(in);
			}
			
			//新建一个sheet工作表
			HSSFSheet sheet = hw.createSheet(Areasheet);
			
			//建立样式
			HSSFCellStyle my_style = hw.createCellStyle();
            // We will now specify a background cell color
			my_style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            my_style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);

            my_style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框    
            my_style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框    
            my_style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框    
            my_style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框   
            
            
			if (sheet.getRow(0)==null){
				
				HSSFRow row = sheet.createRow(0);  
		         //第四步，创建单元格，设置表头  
				HSSFCell cell = row.createCell(0);    
				cell.setCellStyle(my_style);
		         cell.setCellValue("序号");  
		         cell = row.createCell(1);  
		         cell.setCellStyle(my_style);
		         cell.setCellValue("收寄日期");  
		         cell=row.createCell(2);  
		         cell.setCellStyle(my_style);
		         cell.setCellValue("邮件号");  
		         cell=row.createCell(3);  
		         cell.setCellStyle(my_style);
		         cell.setCellValue("寄达省");  
		         cell=row.createCell(4);  
		         cell.setCellStyle(my_style);
		         cell.setCellValue("寄达局");  
		         cell=row.createCell(5);  
		         cell.setCellStyle(my_style);
		         cell.setCellValue("计泡重量");  
		         cell=row.createCell(6);  
		         cell.setCellStyle(my_style);
		         cell.setCellValue("重量");  
		         cell=row.createCell(7);  
		         cell.setCellStyle(my_style);
		         cell.setCellValue("金额");  
		         cell=row.createCell(8);  
		         cell.setCellStyle(my_style);
		         cell.setCellValue("补交");  
		         cell=row.createCell(9);  
		         cell.setCellStyle(my_style);
		         cell.setCellValue("收件人"); 
		         
		         CellRangeAddress c = CellRangeAddress.valueOf("A1:J1");
		         sheet.setAutoFilter(c);
			}
			
		for (int i = 0; i < selectiondata.size(); i++) {
			
			HSSFRow hssfRow = sheet.createRow(i + 1);

			ArrayList<Object> p = selectiondata.get(i);
			hssfRow.createCell(0).setCellValue(p.get(0).toString());
			hssfRow.createCell(1).setCellValue(p.get(1).toString());
			hssfRow.createCell(2).setCellValue(p.get(2).toString());
			hssfRow.createCell(3).setCellValue(p.get(3).toString());
			hssfRow.createCell(4).setCellValue(p.get(4).toString());
			hssfRow.createCell(5).setCellValue((double) p.get(5));
			hssfRow.createCell(6).setCellValue((double) p.get(6));
			hssfRow.createCell(7).setCellValue((double) p.get(7));
			hssfRow.createCell(8).setCellValue((double) p.get(8));
			hssfRow.createCell(9).setCellValue(p.get(9).toString());
				
			
		}
		FileOutputStream fos = new FileOutputStream(fileoutpath);  
        hw.write(fos);  
        System.out.println("恭喜您！写入成功！！！！！！");  
        fos.close();  

	}
	
	
	
	//°指定一行
		public ArrayList<Object> readbyrow(int rowlocal,String selectfilepath)
		{
			ArrayList<Object> a = new ArrayList<Object>();
			
			try {	
				//2011
				//根据绝对路径获得excel文件
				InputStream in = new FileInputStream(selectfilepath);
				//POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("g://hahah.xlsx"));
				//得到excel对象
				HSSFWorkbook hw = new HSSFWorkbook(in);
				//获得第一个工作簿
				HSSFSheet sheet = hw.getSheetAt(0);
						
				//获得sheet之后 对sheet里的每一行进行遍历
				for (int rowindex = rowlocal - 1;rowindex < rowlocal;rowindex++){
					int realrow = rowindex+1;
					HSSFRow hssfRow = sheet.getRow(rowindex);
					if (hssfRow==null){
						continue;
					}
					System.out.println();
					for (int cellindex = hssfRow.getFirstCellNum();cellindex <=hssfRow.getLastCellNum();cellindex++ ){
						Object num = null;
						
						int realcell = cellindex+1;
						HSSFCell cell = hssfRow.getCell(cellindex);

						if (cell == null) {
							continue;
						} else if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
							continue;
						} else if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
							num = cell.getRichStringCellValue().getString();
							//System.out.print("第"+realrow+"行 第"+realcell+"列"+cell.getRichStringCellValue().getString()+"     ");
						} else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
							num = (double) cell.getNumericCellValue();
						   // System.out.print("第"+realrow+"行 第"+realcell+"列"+cell.getNumericCellValue()+"     ");
						} else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
						} else if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
						}
						
						a.add(num);
					}
					
				}
				
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				return null;
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return a;
			
		}
	
	
	//打印指定一行
	public void print(int rowlocal)
	{
		try {	
			//2011
			//根据绝对路径获得excel文件
			InputStream in = new FileInputStream("g://222.xls");
			//POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("g://hahah.xlsx"));
			//得到excel对象
			HSSFWorkbook hw = new HSSFWorkbook(in);
			//获得第一个工作簿
			HSSFSheet sheet = hw.getSheetAt(0);
					
			//获得sheet之后 对sheet里的每一行进行遍历
			for (int rowindex = rowlocal - 1;rowindex < rowlocal;rowindex++){
				int realrow = rowindex+1;
				HSSFRow hssfRow = sheet.getRow(rowindex);
				if (hssfRow==null){
					continue;
				}
				System.out.println();
				for (int cellindex = hssfRow.getFirstCellNum();cellindex <=hssfRow.getLastCellNum();cellindex++ ){
					int realcell = cellindex+1;
					HSSFCell cell = hssfRow.getCell(cellindex);

					if (cell == null) {
						continue;
					} else if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
						continue;
					} else if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
						System.out.print("第"+realrow+"行 第"+realcell+"列"+cell.getRichStringCellValue().getString()+"     ");
					} else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
					    System.out.print("第"+realrow+"行 第"+realcell+"列"+cell.getNumericCellValue()+"     ");
					} else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
					} else if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
					}
				}
				
			}
			
			
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
}
