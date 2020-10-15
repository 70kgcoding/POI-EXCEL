package com.wdh.exceldemo.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ForTeS1 {

	public static void excelRead() throws Exception {
		//获得Excel文件输出流

		InputStream in = new FileInputStream("D:\\tmp/test2.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //遍历行
        String tag = "asdasdas";
		int k = 1;
    	for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {

    		XSSFRow row = sheet.getRow(i);
    		String miaoshu=row.getCell(0).getStringCellValue();
    		if(tag.equals(miaoshu))
    		{
    			k++;
    			row.createCell(2).setCellValue(k);
    		}
    		else
    		{
    			k = 1;
    			row.createCell(2).setCellValue(k);
    			tag = miaoshu;
    			
    		}
    			
    			
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
