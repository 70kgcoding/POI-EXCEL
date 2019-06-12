package com.wdh.exceldemo.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Hashtable;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestForhuaxuepinlaiyuan {
	private static Hashtable<String, String> pool=new Hashtable<String, String>();
	static {
		pool.put("MFBDT2","_SDS_DH_36");
		pool.put("MFPA2","_SDS_DH_37");
		pool.put("MFLL2","_SDS_DH_38");
		pool.put("MFEOEG2","_SDS_DH_39");
		pool.put("MFHD2","_SDS_DH_40");
		pool.put("MFPP2","_SDS_DH_41");
		pool.put("MFLOP2","_SDS_DH_42");
		pool.put("MFHDPE2/LLDPE2","_SDS_DH_43");
		pool.put("MFLLDPE2","_SDS_DH_44");
		pool.put("LOP2","_SDS_DH_45");
		pool.put("OM2","_SDS_DH_46");
		pool.put("MFLO2","_SDS_DH_47");
		pool.put("MFOM2","_SDS_DH_48");
		pool.put("MFUT2","_SDS_DH_49");
		pool.put("中海石油炼化惠州炼油分公司","_SDS_DH_50");
		pool.put("惠州石化","_SDS_DH_51");
		pool.put("丁二烯抽提","_SDS_DH_52");
		pool.put("PA","_SDS_DH_53");
		pool.put("EOEG","_SDS_DH_54");
		pool.put("OXO","_SDS_DH_55");
		pool.put("OM","_SDS_DH_56");
		pool.put("供应商","_SDS_DH_26");
		pool.put("普莱克斯","_SDS_DH_33");
		pool.put("MFSP","_SDS_DH_19");
		pool.put("MFLO","_SDS_DH_15");
		pool.put("艾尔特","_SDS_DH_28");
		pool.put("空气化工","_SDS_DH_31");
		pool.put("Bolinda"," _SDS_DH_3");
		pool.put("Merck","_SDS_DH_13");
		pool.put("Hach","_SDS_DH_11");
	}

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
	    		if(pool.keySet().contains(miaoshu))
	    			row.createCell(0).setCellValue(pool.get(miaoshu));	
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
