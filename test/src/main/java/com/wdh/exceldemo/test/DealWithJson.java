package com.wdh.exceldemo.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Hashtable;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DealWithJson {
	private static Hashtable<String, String> pool=new Hashtable<String, String>();
	static {
		pool.put("北京市","_SYS_A1_401");
		pool.put("天津市","_SYS_A1_402");
		pool.put("河北省","_SYS_A1_403");
		pool.put("山西省","_SYS_A1_404");
		pool.put("内蒙古自治区","_SYS_A1_405");
		pool.put("辽宁省","_SYS_A1_406");
		pool.put("吉林省","_SYS_A1_407");
		pool.put("黑龙江省","_SYS_A1_408");
		pool.put("上海市","_SYS_A1_409");
		pool.put("江苏省","_SYS_A1_410");
		pool.put("浙江省","_SYS_A1_411");
		pool.put("安徽省","_SYS_A1_412");
		pool.put("福建省","_SYS_A1_413");
		pool.put("江西省","_SYS_A1_414");
		pool.put("山东省","_SYS_A1_415");
		pool.put("河南省","_SYS_A1_416");
		pool.put("湖北省","_SYS_A1_417");
		pool.put("湖南省","_SYS_A1_418");
		pool.put("广东省","_SYS_A1_419");
		pool.put("广西壮族自治区","_SYS_A1_420");
		pool.put("海南省","_SYS_A1_421");
		pool.put("重庆市","_SYS_A1_422");
		pool.put("四川省","_SYS_A1_423");
		pool.put("贵州省","_SYS_A1_424");
		pool.put("云南省","_SYS_A1_425");
		pool.put("西藏自治区","_SYS_A1_426");
		pool.put("陕西省","_SYS_A1_427");
		pool.put("甘肃省","_SYS_A1_428");
		pool.put("青海省","_SYS_A1_429");
		pool.put("宁夏回族自治区","_SYS_A1_430");
		pool.put("新疆维吾尔自治区","_SYS_A1_431");
		pool.put("台湾省","_SYS_A1_432");
		pool.put("香港特别行政区","_SYS_A1_433");
		pool.put("澳门特别行政区","_SYS_A1_434");
	}
	public static void excelRead() throws Exception {
		//获得Excel文件输出流

		InputStream in = new FileInputStream("D:\\tmp/testforjson.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //遍历行
    	for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
    		XSSFRow row = sheet.getRow(i);
    		if (row.getCell(1)!= null && row.getCell(1).getStringCellValue() != null) {
    			row.getCell(1).setCellValue(pool.get(row.getCell(1).getStringCellValue()));
    		}
    			
		}	
    	OutputStream out = new FileOutputStream("D:\\tmp/testforjson.xlsx");
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
