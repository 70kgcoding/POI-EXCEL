package com.wdh.exceldemo.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Hashtable;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
   * 二期化学品
 *
 */
public class App 
{
	
	private static Hashtable<String, Integer> pool=new Hashtable<String, Integer>();
	static {
		pool.put("_SDS_AS_1",1);
		pool.put("_SDS_AT_1",2);
		pool.put("_SDS_AU_1",3);
		pool.put("_SDS_AV_1",4);
		pool.put("_SDS_AW_1",5);
		pool.put("_SDS_AX_1",6);
		pool.put("_SDS_AY_1",7);
		pool.put("_SDS_AZ_1",8);
		pool.put("_SDS_BA_1",9);
		pool.put("_SDS_BB_1",10);
		pool.put("_SDS_BC_1",11);
		pool.put("_SDS_BD_1",12);
		pool.put("_SDS_BE_1",13);
		pool.put("_SDS_BF_1",14);
		pool.put("_SDS_BG_1",15);
		pool.put("_SDS_BH_1",16);
		pool.put("_SDS_BI_1",17);
		pool.put("_SDS_BJ_1",18);
		pool.put("_SDS_BK_1",19);
		pool.put("_SDS_BL_1",20);
		pool.put("_SDS_BM_1",21);
		pool.put("_SDS_BN_1",22);
		pool.put("_SDS_BO_1",23);
		pool.put("_SDS_BP_1",24);
		pool.put("_SDS_BQ_1",25);
		pool.put("_SDS_CB_1",26);
		pool.put("_SDS_BS_1",27);
		pool.put("_SDS_BT_1",28);
		pool.put("_SDS_BU_1",29);
		pool.put("_SDS_BV_1",30);
		pool.put("_SDS_BW_1",31);
		pool.put("_SDS_BX_1",32);
		pool.put("_SDS_BY_1",33);
		pool.put("_SDS_BZ_1",34);
		pool.put("_SDS_CA_1",35);
	}
	public static void excelWrite() throws Exception {
		//获得Excel文件输出流

		OutputStream out = new FileOutputStream("D:\\tmp/test.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();
//        样式
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);


        XSSFSheet sheet = workbook.createSheet("01");
//        rows
        XSSFRow row = sheet.createRow(0);
//        cells
        XSSFCell c1 = row.createCell(0);

        c1.setCellValue("dada");
        XSSFCell c2 = row.createCell(1);
        c2.setCellValue("工资");
        workbook.write(out);
        out.close();
	}

	
	
	public static void excelRead() throws Exception {
		//获得Excel文件输出流

		InputStream in = new FileInputStream("D:\\tmp/test.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //遍历行
    	for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
    		XSSFRow row = sheet.getRow(i);
    		String miaoshu=row.getCell(0).getStringCellValue();
    		String des[]=miaoshu.split(",");
    		for(int j=0;j<des.length;j++)
    		{
    			String temp=des[j].substring(0, des[j].length()-1)+"1";
    			if(pool.keySet().contains(temp))
    			{
    				row.createCell(pool.get(temp)).setCellValue(des[j]);
    			}
    		}
		}	
    	OutputStream out = new FileOutputStream("D:\\tmp/test.xlsx");
    	workbook.write(out);
        out.close();
	}

	
    public static void main( String[] args )
    {
    	try {
    		excelRead();
    		System.out.println("success");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    	
    }
}
