package com.wdh.exceldemo.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 处理行业类别 多级类别数据导入
 * @author Administrator
 *
 */
public class DealWithEPMS1755 {

	public DealWithEPMS1755() {
		// TODO Auto-generated constructor stub
	}

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		InputStream in = new FileInputStream("D:\\tmp/test0726.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //遍历行
        // 门类
        String parent1 = "";
        // 大类
        double parent2 = 0;
        
        // 
        String parent3 = "";
    	for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
    		XSSFRow row = sheet.getRow(i);
    		if (row.getCell(0).getCellTypeEnum().equals(CellType.BLANK)) {
    			row.createCell(0).setCellValue(parent2);    		
    		} else {
    			parent2 = row.getCell(0).getNumericCellValue();
    			int e =0;
    		}
    	}
    	OutputStream out = new FileOutputStream("D:\\tmp/test0726.xlsx");
    	workbook.write(out);
        out.close();
        System.out.println("excute success");
	}
}
