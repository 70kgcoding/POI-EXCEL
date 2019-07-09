package com.wdh.exceldemo.test;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

public class Json2Excel {
	public static String readJsonData(String pactFile) throws IOException {
		// 读取文件数据
		//System.out.println("读取文件数据util");
		
		StringBuffer strbuffer = new StringBuffer();
		File myFile = new File(pactFile);//"D:"+File.separatorChar+"DStores.json"
		if (!myFile.exists()) {
			System.err.println("Can't Find " + pactFile);
		}
		try {
			FileInputStream fis = new FileInputStream(pactFile);
			InputStreamReader inputStreamReader = new InputStreamReader(fis, "UTF-8");
			BufferedReader in  = new BufferedReader(inputStreamReader);
			
			String str;
			while ((str = in.readLine()) != null) {
				strbuffer.append(str);  //new String(str,"UTF-8")
			}
			in.close();
		} catch (IOException e) {
			e.getStackTrace();
		}
		//System.out.println("读取文件结束util");
		return strbuffer.toString();
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			InputStream in = new FileInputStream("D:\\tmp/testforjson.xlsx");
	        XSSFWorkbook workbook = new XSSFWorkbook(in);
	        XSSFSheet sheet = workbook.getSheetAt(0);
			
			String json = readJsonData("C:\\Users\\Administrator\\Desktop\\China.json");
	        JSONObject jsonObject = JSONObject.parseObject(json);
	        //获取到国家及下面所有的信息 开始循环插入，这里可以写成递归调用，但是不如这样方便查看、理解
	        JSONArray countryAll = jsonObject.getJSONArray("districts");
	        int tag = 1;
	        for (int i = 0; i < countryAll.size(); i++) {
	            JSONObject countryLeve0 = countryAll.getJSONObject(i);
	            String citycode0 = countryLeve0.getString("citycode");
	            String adcode0 = countryLeve0.getString("adcode");
	            String name0 = countryLeve0.getString("name");
	            String center0 = countryLeve0.getString("center");
	            String country = countryLeve0.getString("level");
	            int level = 0;
	            if (country.equals("country")) {
	                level = 0;
	            }
	            //插入国家
                sheet.getRow(tag).createCell(0).setCellValue(name0);
	            System.out.println(tag);
	            JSONArray province0 = countryLeve0.getJSONArray("districts");

	            for (int j = 0; j < province0.size(); j++) {
	                JSONObject province1 = province0.getJSONObject(j);
	                String citycode1 = province1.getString("citycode");
	                String adcode1 = province1.getString("adcode");
	                String name1 = province1.getString("name");
	                String center1 = province1.getString("center");
	                String province = province1.getString("level");
	                int level1 = 0;
	                if (province.equals("province")) {
	                    level1 = 1;
	                }
	                int tag4city = 1;
	                tag++;
	                //插入省
                    sheet.getRow(tag).createCell(1).setCellValue(name1);
	                System.out.println(tag);
	                JSONArray city0 = province1.getJSONArray("districts");

	                for (int z = 0; z < city0.size(); z++) {
	                    JSONObject city2 = city0.getJSONObject(z);
	                    String citycode2 = city2.getString("citycode");
	                    String adcode2 = city2.getString("adcode");
	                    String name2 = city2.getString("name");
	                    String center2 = city2.getString("center");
	                    String city = city2.getString("level");
	                    int level2 = 0;
	                    if (city.equals("city")) {
	                        level2 = 2;
	                    }
	                    int tag4street = 1;
	                    tag++;
	                    //插入市
	                    sheet.getRow(tag).createCell(2).setCellValue(name2);
	                    //sheet.getRow(tag).createCell(1).setCellValue(name1);
	                    sheet.getRow(tag).createCell(4).setCellValue(tag4city);
	                    System.out.println(tag);
	                    tag4city++;

	                    JSONArray street0 = city2.getJSONArray("districts");
	                    for (int w = 0; w < street0.size(); w++) {
	                        JSONObject street3 = street0.getJSONObject(w);
	                        String citycode3 = street3.getString("citycode");
	                        String adcode3 = street3.getString("adcode");
	                        String name3 = street3.getString("name");
	                        String center3 = street3.getString("center");
	                        String street = street3.getString("level");
	                        int level3 = 0;
	                        if (street.equals("street")) {
	                            level3 = 2;
	                        }
	                        tag++;
	                        //插入区县
	                        //sheet.getRow(tag).createCell(1).setCellValue(name1);
	                        //sheet.getRow(tag).createCell(2).setCellValue(name2);
	                        sheet.getRow(tag).createCell(3).setCellValue(name3);
	                        sheet.getRow(tag).createCell(4).setCellValue(tag4street);
	                        tag4street++;
	                        System.out.println(tag);	                        //  JSONArray street = street3.getJSONArray("districts");
	                        //有需要可以继续向下遍历

	                    }

	                }
	            }
	            tag++;
	        }
	        OutputStream out = new FileOutputStream("D:\\tmp\\testforjson.xlsx");
	    	workbook.write(out);
	        out.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
