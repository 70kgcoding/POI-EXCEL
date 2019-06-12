package com.wdh.exceldemo.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ForfujianType {

	public static void excelRead() throws Exception {
		//获得Excel文件输出流

		InputStream in = new FileInputStream("D:\\tmp/test2.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //遍历行
    	for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
    		XSSFRow row = sheet.getRow(i);
    		String miaoshu=row.getCell(0).getStringCellValue();
    		if(miaoshu==null)
    			continue;
    		if(miaoshu.substring(miaoshu.length()-3).equals("pdf")||miaoshu.substring(miaoshu.length()-3).equals("PDF"))
    			row.createCell(1).setCellValue("pdf");
    		else
    			row.createCell(1).setCellValue("doc");
    			
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
