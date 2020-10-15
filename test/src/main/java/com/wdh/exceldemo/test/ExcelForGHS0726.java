package com.wdh.exceldemo.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelForGHS0726 {

	
	
	public static void excute() throws Exception {
		
		String[] ghs01 = {"_SDS_AY_2","_SDS_AY_3","_SDS_BD_2","_SDS_BD_3","_SDS_BU_2","_SDS_BU_3","_SDS_BU_4","_SDS_BU_5","_SDS_BU_6"};
		String[] ghs02 = {"_SDS_AS_2","_SDS_AT_2","_SDS_AT_3","_SDS_AW_2","_SDS_AW_3","_SDS_AW_4","_SDS_AX_2","_SDS_AX_3","_SDS_AY_4","_SDS_AY_5","_SDS_AY_6","_SDS_AY_7","_SDS_AZ_2","_SDS_AZ_3","_SDS_BA_2","_SDS_BA_3","_SDS_BA_4","_SDS_BD_3","_SDS_BD_4","_SDS_BD_5","_SDS_BD_6","_SDS_BD_7","_SDS_BV_2","_SDS_BW_2"};
		String[] ghs03 = {"_SDS_AU_2","_SDS_BB_2","_SDS_BB_3","_SDS_BB_4","_SDS_BC_2","_SDS_BC_3","_SDS_BC_4"};
		String[] ghs04 = {"_SDS_AV_3","_SDS_AV_2"};
		String[] ghs05 = {"_SDS_BE_2","_SDS_BE_3","_SDS_BE_4","_SDS_BI_2","_SDS_BI_3","_SDS_BI_4","_SDS_BI_5","_SDS_BI_8","_SDS_BJ_2"};
		String[] ghs06 = {"_SDS_BF_2","_SDS_BF_3","_SDS_BF_4","_SDS_BG_2","_SDS_BG_3","_SDS_BG_4","_SDS_BH_2","_SDS_BH_3"};
		String[] ghs07 = {"_SDS_BF_5","_SDS_BG_5","_SDS_BH_5","_SDS_BQ_4","_SDS_BQ_5","_SDS_CB_4","_SDS_CB_5"};
		String[] ghs07Special = {"_SDS_BI_6","_SDS_BJ_3","_SDS_BJ_4","_SDS_BJ_5","_SDS_BJ_8","_SDS_BJ_9","_SDS_BK_2","_SDS_BK_3","_SDS_BK_4","_SDS_BK_5"};
		String[] ghs08 = {"_SDS_BL_2","_SDS_BL_3","_SDS_BM_2","_SDS_BM_3","_SDS_BM_4","_SDS_BN_2","_SDS_BN_3","_SDS_BN_4","_SDS_BN_5","_SDS_BN_6","_SDS_BO_10","_SDS_BO_11","_SDS_BO_12","_SDS_BO_13","_SDS_BO_14","_SDS_BO_15","_SDS_BO_16","_SDS_BO_17","_SDS_BO_18","_SDS_BO_19","_SDS_BO_2","_SDS_BO_3","_SDS_BO_4","_SDS_BO_5","_SDS_BO_6","_SDS_BO_7","_SDS_BO_8","_SDS_BO_9","_SDS_BQ_2","_SDS_BQ_3","_SDS_CB_2","_SDS_CB_3","_SDS_BS_2","_SDS_BS_3","_SDS_BT_2","_SDS_BT_3"};
		String[] ghs09 = {"_SDS_BX_2","_SDS_BY_2","_SDS_BY_3"};
		String[] warning = {"_SDS_AS_3","_SDS_AT_3","_SDS_AV_2","_SDS_AV_3","_SDS_AW_4","_SDS_AW_5","_SDS_AX_3","_SDS_AY_6","_SDS_AY_7","_SDS_AZ_3","_SDS_BA_4","_SDS_BB_4","_SDS_BC_4","_SDS_BD_6","_SDS_BD_7","_SDS_BE_2","_SDS_BE_3","_SDS_BE_4","_SDS_BF_5","_SDS_BF_6","_SDS_BG_5","_SDS_BG_6","_SDS_BH_5","_SDS_BH_6","_SDS_BI_6","_SDS_BI_7","_SDS_BJ_3","_SDS_BJ_4","_SDS_BJ_5","_SDS_BJ_6","_SDS_BJ_8","_SDS_BJ_9","_SDS_BK_2","_SDS_BK_3","_SDS_BK_4","_SDS_BK_5","_SDS_BM_4","_SDS_BN_6","_SDS_BO_14","_SDS_BO_15","_SDS_BO_16","_SDS_BO_17","_SDS_BO_18","_SDS_BO_19","_SDS_BQ_3","_SDS_BQ_4","_SDS_BQ_5","_SDS_CB_3","_SDS_CB_4","_SDS_CB_5","_SDS_BS_2","_SDS_BT_3","_SDS_BU_6","_SDS_BX_2","_SDS_BY_2"};
		String[] danger = {"_SDS_AS_2","_SDS_AT_2","_SDS_AU_2","_SDS_AW_2","_SDS_AW_3","_SDS_AX_2","_SDS_AY_2","_SDS_AY_3","_SDS_AY_4","_SDS_AY_5","_SDS_AZ_2","_SDS_BA_2","_SDS_BA_3","_SDS_BB_2","_SDS_BB_3","_SDS_BC_2","_SDS_BC_3","_SDS_BD_2","_SDS_BD_3","_SDS_BD_4","_SDS_BD_5","_SDS_BF_2","_SDS_BF_3","_SDS_BF_4","_SDS_BG_2","_SDS_BG_3","_SDS_BG_4","_SDS_BH_2","_SDS_BH_3","_SDS_BH_4","_SDS_BI_2","_SDS_BI_3","_SDS_BI_4","_SDS_BI_5","_SDS_BI_8","_SDS_BJ_2","_SDS_BJ_7","_SDS_BL_2","_SDS_BL_3","_SDS_BM_2","_SDS_BM_3","_SDS_BN_2","_SDS_BN_3","_SDS_BN_4","_SDS_BN_5","_SDS_BO_10","_SDS_BO_11","_SDS_BO_12","_SDS_BO_13","_SDS_BO_2","_SDS_BO_3","_SDS_BO_4","_SDS_BO_5","_SDS_BO_6","_SDS_BO_7","_SDS_BO_8","_SDS_BO_9","_SDS_BQ_2","_SDS_CB_2","_SDS_BS_3","_SDS_BT_2","_SDS_BU_2","_SDS_BU_3","_SDS_BU_4","_SDS_BU_5","_SDS_BU_7","_SDS_BV_2","_SDS_BW_2"};
		
		ArrayList<String> ghs01List = new ArrayList<String>(Arrays.asList(ghs01));
		ArrayList<String> ghs02List = new ArrayList<String>(Arrays.asList(ghs02));
		ArrayList<String> ghs03List = new ArrayList<String>(Arrays.asList(ghs03));
		ArrayList<String> ghs04List = new ArrayList<String>(Arrays.asList(ghs04));
		ArrayList<String> ghs05List = new ArrayList<String>(Arrays.asList(ghs05));
		ArrayList<String> ghs06List = new ArrayList<String>(Arrays.asList(ghs06));
		ArrayList<String> ghs07List = new ArrayList<String>(Arrays.asList(ghs07));
		ArrayList<String> ghs07SpecialList = new ArrayList<String>(Arrays.asList(ghs07Special));
		ArrayList<String> ghs08List = new ArrayList<String>(Arrays.asList(ghs08));
		ArrayList<String> ghs09List = new ArrayList<String>(Arrays.asList(ghs09));
		ArrayList<String> warningList = new ArrayList<String>(Arrays.asList(warning));
		ArrayList<String> dangerList = new ArrayList<String>(Arrays.asList(danger));
		
		//获得Excel文件输出流
		InputStream in = new FileInputStream("D:\\tmp/test0726.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //遍历行
    	for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
    		XSSFRow row = sheet.getRow(i);
    		String category=row.getCell(1).getStringCellValue();
    		String[] categories= category.split(",");
    		ArrayList<String> categoriesList = new ArrayList<String>(Arrays.asList(categories)); 
    		for (int j = 0;j < categories.length;j++) {
    			if (ghs01List.contains(categories[j])) {
    				row.createCell(2).setCellValue("1");
    			}
    			if (ghs02List.contains(categories[j])) {
    				row.createCell(3).setCellValue("1");
    			}
    			if (ghs03List.contains(categories[j])) {
    				row.createCell(4).setCellValue("1");
    			}
    			if (ghs04List.contains(categories[j])) {
    				row.createCell(5).setCellValue("1");
    			}
    			if (ghs05List.contains(categories[j])) {
    				row.createCell(6).setCellValue("1");
    			}
    			if (ghs06List.contains(categories[j])) {
    				row.createCell(7).setCellValue("1");
    			}
    			if (ghs07List.contains(categories[j])) {
    				row.createCell(8).setCellValue("1");
    			}
    			if (ghs08List.contains(categories[j])) {
    				row.createCell(9).setCellValue("1");
    			}
    			if (ghs09List.contains(categories[j])) {
    				row.createCell(10).setCellValue("1");
    			}	
    			if (row.getCell(11) != null && row.getCell(11).getStringCellValue().equals("_SDS_AH_3")) {
    				// do nothing
    			}else {
        			if (warningList.contains(categories[j])) {
        				row.createCell(11).setCellValue("_SDS_AH_2");
        			}
        			if (dangerList.contains(categories[j])) {
        				row.createCell(11).setCellValue("_SDS_AH_3");
        			}
    			}
                if (!categoriesList.contains("_SDS_BL_2") && ghs07SpecialList.contains(categories[j])) {
                	row.createCell(8).setCellValue("1");
    			}
    		}
		}	
    	OutputStream out = new FileOutputStream("D:\\tmp/test0726.xlsx");
    	workbook.write(out);
        out.close();
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		try {
			excute();
			System.out.println("excute success");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
		
	}

}
