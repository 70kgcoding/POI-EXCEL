package com.wdh.exceldemo.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Hashtable;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ForTeS2 {

	private static Hashtable<String, String> map=new Hashtable<String, String>();
	static {
		map.put("HW01医疗废物","EPMS_COMMON_A2_1");
		map.put("HW02医药废物","EPMS_COMMON_A2_2");
		map.put("HW03废药物、药品","EPMS_COMMON_A2_3");
		map.put("HW04农药废物","EPMS_COMMON_A2_4");
		map.put("HW05木材防腐剂废物","EPMS_COMMON_A2_5");
		map.put("HW06废有机溶剂与含有机溶剂 废物","EPMS_COMMON_A2_6");
		map.put("HW07热处理含氰废物","EPMS_COMMON_A2_7");
		map.put("HW08废矿物油与含矿物油废物","EPMS_COMMON_A2_8");
		map.put("HW09 油/水、烃/水混合物或乳化液","EPMS_COMMON_A2_9");
		map.put("HW10多氯（溴）联苯类 废物","EPMS_COMMON_A2_10");
		map.put("HW11精（蒸）馏残渣","EPMS_COMMON_A2_11");
		map.put("HW12染料、涂料废物","EPMS_COMMON_A2_12");
		map.put("HW13有机树脂类废物","EPMS_COMMON_A2_13");
		map.put("HW14新化学物质废物","EPMS_COMMON_A2_14");
		map.put("HW15爆炸性废物","EPMS_COMMON_A2_15");
		map.put("HW16感光材料废物","EPMS_COMMON_A2_16");
		map.put("HW17表面处理废物","EPMS_COMMON_A2_17");
		map.put("HW18焚烧处置残渣","EPMS_COMMON_A2_18");
		map.put("HW19含金属羰基化合物废物","EPMS_COMMON_A2_19");
		map.put("HW20含铍废物","EPMS_COMMON_A2_20");
		map.put("HW21含铬废物","EPMS_COMMON_A2_21");
		map.put("HW22含铜废物","EPMS_COMMON_A2_22");
		map.put("HW23含锌废物","EPMS_COMMON_A2_23");
		map.put("HW24含砷废物","EPMS_COMMON_A2_24");
		map.put("HW25含硒废物","EPMS_COMMON_A2_25");
		map.put("HW26含镉废物","EPMS_COMMON_A2_26");
		map.put("HW27含锑废物","EPMS_COMMON_A2_27");
		map.put("HW28含碲废物","EPMS_COMMON_A2_28");
		map.put("HW29含汞废物","EPMS_COMMON_A2_29");
		map.put("HW30含铊废物","EPMS_COMMON_A2_30");
		map.put("HW31含铅废物","EPMS_COMMON_A2_31");
		map.put("HW32无机氟化物废物","EPMS_COMMON_A2_32");
		map.put("HW33无机氰化物废物","EPMS_COMMON_A2_33");
		map.put("HW34废酸","EPMS_COMMON_A2_34");
		map.put("HW35废碱","EPMS_COMMON_A2_35");
		map.put("HW36石棉废物","EPMS_COMMON_A2_36");
		map.put("HW37有机磷化合物废物","EPMS_COMMON_A2_37");
		map.put("HW38有机氰化物废物","EPMS_COMMON_A2_38");
		map.put("HW39含酚废物","EPMS_COMMON_A2_39");
		map.put("HW40含醚废物","EPMS_COMMON_A2_40");
		map.put("HW45含有机卤化物废物","EPMS_COMMON_A2_41");
		map.put("HW46含镍废物","EPMS_COMMON_A2_42");
		map.put("HW47含钡废物","EPMS_COMMON_A2_43");
		map.put("HW48有色金属冶炼废物","EPMS_COMMON_A2_44");
		map.put("HW49其他废物","EPMS_COMMON_A2_45");
		map.put("HW50废催化剂","EPMS_COMMON_A2_46");
	}
	
	public static void excelRead() throws Exception {
		//获得Excel文件输出流

		InputStream in = new FileInputStream("D:\\tmp/test2.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //遍历行
	
    	for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {

    		XSSFRow row = sheet.getRow(i);
    		String miaoshu=row.getCell(0).getStringCellValue();
    		row.createCell(3).setCellValue(map.get(miaoshu));
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
