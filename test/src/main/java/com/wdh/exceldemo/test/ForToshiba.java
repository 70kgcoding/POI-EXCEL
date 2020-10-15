package com.wdh.exceldemo.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Hashtable;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ForToshiba {

		public static void excelRead() throws Exception {
			//获得Excel文件输出流

			InputStream in = new FileInputStream("D:\\tmp/test2.xlsx");
	        XSSFWorkbook workbook = new XSSFWorkbook(in);
	        XSSFSheet sheet = workbook.getSheetAt(0);
	        //遍历行
	        double temp = -9999;
	        int tag = 1;
	    	for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
	    		XSSFRow row = sheet.getRow(i);
	    		double seq=row.getCell(0).getNumericCellValue();
	    		if (temp == seq) {
	    			row.createCell(2).setCellValue(++tag);
	    		} else {
	    			tag = 1;
	    			row.createCell(2).setCellValue(1);
	    		}
	    		temp = seq;

			}	
	    	OutputStream out = new FileOutputStream("D:\\tmp/test2.xlsx");
	    	workbook.write(out);
	        out.close();
		}
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			excelRead();
			System.out.println("success");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}
