package com.wdh.exceldemo.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Hashtable;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestForzhizaoshang {
	private static Hashtable<String, String> pool=new Hashtable<String, String>();
	static {
		pool.put("中海壳牌石油化工有限公司","MF-0003");
		pool.put("Conqord Oil","MF-0009");
		pool.put("湖北华邦化学有限公司","MF-0010");
		pool.put("THE DOW CHEMICAL COMPANY","MF-0028");
		pool.put("壳牌(中国)有限公司","MF-0033");
		pool.put("通用电气水处理技术（无锡）有限公司","MF-0039");
		pool.put("空气化工产品（上海）有限公司","MF-0040");
		pool.put("纳尔科（中国）环保技术服务有限公司","MF-0041");
		pool.put("扬子石化—巴斯夫有限责任公司","MF-0045");
		pool.put("Mitsubishi Chemical Corporation","MF-0049");
		pool.put("新浦化学（泰兴）有限公司","MF-0050");
		pool.put("空气化工产品气体（深圳）有限公司","MF-0052");
		pool.put("CNOOC and Shell Petrochemicals Company Limited","MF-0060");
		pool.put("惠州联宏化工有限公司","MF-0068");
		pool.put("东营石大宏益化工有限公司","MF-0082");
		pool.put("DOW CHEMICAL PACIFIC LIMITED","MF-0088");
		pool.put("上海环球分子筛有限公司","MF-0092");
		pool.put("UOP LLC","MF-0093");
		pool.put("Clariant Produkte (Deutschland) GmbH","MF-0097");
		pool.put("惠州市深华化工有限公司","MF-0121");
		pool.put("惠州市百利宏晟安化工有限公司","MF-0122");
		pool.put("普莱克斯（惠州）工业气体有限公司","MF-0126");
		pool.put("国药集团化学试剂有限公司","MF-0144");
		pool.put("西格玛奥德里奇（上海）贸易有限公司","MF-0145");
		pool.put("Acros Organics BVBA","MF-0146");
		pool.put("广州化学试剂厂","MF-0149");
		pool.put("Thermo Fisher Scientific","MF-0151");
		pool.put("深圳市博林达科技有限公司","MF-0153");
		pool.put("默克股份两合公司","MF-0154");
		pool.put("Hach Company","MF-0156");
		pool.put("营口市风光化工有限公司","MF-0168");
		pool.put("东莞市汉维新材料科技有限公司","MF-0170");
		pool.put("纪州喷码技术（上海）有限公司","MF-0177");
		pool.put("尤尼维讯（张家港）化学有限公司","MF-0178");
		pool.put("上海罗门哈斯化工有限公司","MF-0179");
		pool.put("ADSORBENTS,CATALYSTS&SERVICES","MF-0180");
		pool.put("深圳市鸿瑞化工工程有限公司","MF-0181");
		pool.put("木林森活性炭江苏有限公司","MF-0182");
		pool.put("惠州市易普达机电设备有限公司","MF-0183");
		pool.put("中山市华明泰化工股份有限公司","MF-0184");
		pool.put("营口风光新材料股份有限公司","MF-0185");
		pool.put("辽宁鼎际得石化股份有限公司","MF-0186");
		pool.put("三菱化学株式会社","MF-0187");
		pool.put("巴氏夫应用化工有限公司","MF-0188");
		pool.put("新兴化学工业株式会社","MF-0189");
		pool.put("凯美特","MF-0190");
		pool.put("加油站","MF-0191");
		pool.put("普莱克斯","MF-0192");
		pool.put("天津林圣金海化工有限公司","MF-0193");
		pool.put("东营市俊源石油技术开发有限公司","MF-0194");
		pool.put("INEOS Polyolefins","MF-0195");
		pool.put("PQ Corporation.","MF-0196");
		pool.put("SASOL Germany GmbH Geschäftsbereich Tenside Paul-Baumann-Str.145764 Marl","MF-0197");
		pool.put("纳尔科(中国)环保技术服务有限公司","MF-0198");
		pool.put("MAGNAPORE®963","MF-0199");
		pool.put("PCC Chemax, Inc","MF-0200");
		pool.put("Akzo Nobel Functional Chemicals B.V.","MF-0201");
		pool.put("納爾科股份有限公司","MF-0202");
		pool.put("ECOLAB PTE LTD","MF-0203");
		pool.put("INEOS Sales (UK) Ltd","MF-0204");
		pool.put("中国石油化工股份有限公司北京燕山分公司","MF-0205");
		pool.put("巴斯夫新材料有限公司 ","MF-0206");
		pool.put("BASF SE","MF-0207");
		pool.put("United Initiators GmbH & Co. KG","MF-0208");
		pool.put("TOTAL LUBRICANTS CHINA CO. LTD.","MF-0209");
		pool.put("北京百灵威科技有限公司","MF-0210");
		pool.put("阿法埃沙（中国）化学有限公司","MF-0211");
		pool.put("SHELLEASTERN CHEMICALS(S)A REGISTERED BUSINESS OF SHELL EASTERN TRADING (PTE) LTD","MF-0212");
		pool.put("Fisher Scientific","MF-0213");
		pool.put("上海麦克林生化科技有限公司","MF-0214");
		pool.put("梯希爱（上海）化成工业发展有限公司","MF-0215");
		pool.put("上海韶远试剂有限公司","MF-0216");
		pool.put("Sciencelab.com, Inc.","MF-0217");
		pool.put("HACH CHINA23","MF-0218");
		pool.put("HACH CHINA24","MF-0219");
		pool.put("HACH CHINA25","MF-0220");
		pool.put("HACH CHINA26","MF-0221");
		pool.put("HORIBA, Ltd.","MF-0222");
		pool.put("天津市科密欧化学试剂有限公司","MF-0223");
		pool.put("Merck","MF-0224");
		pool.put("江苏强盛功能化学股份有限公司","MF-0225");
		pool.put("Ark Pharm, Inc.","MF-0226");
		pool.put("尤尼维讯(张家港)化学有限公司","MF-0227");
		pool.put("DOW CHEMICAL (CHINA) INVESTMENT","MF-0228");
		pool.put("PMC Biogenix, Inc.","MF-0229");
		pool.put("呈和科技股份有限公司","MF-0230");
		pool.put("BASF Advanced Chemicals Co., Ltd","MF-0231");
		pool.put("UOP  LLC","MF-0232");
		pool.put("UNIVATION TECHNOLOGIES, LLC","MF-0233");
		pool.put("中海油惠州石化有限公司","MF-0234");
		pool.put("江苏太湖新材料控股有限公司","MF-0235");
		pool.put("巴斯夫应用化工有限公司中国","MF-0236");
		pool.put("Nalco Australia","MF-0237");
		pool.put("THIOCHIMIE","MF-0238");
		pool.put("中国石化集团茂名石化乙烯工业公司","MF-0239");
		pool.put("ATOFINA","MF-0240");
		pool.put("广州赫尔普化工有限公司","MF-0241");
		pool.put("Johnson Matthey Plc,","MF-0242");
		pool.put("张家港北兴化工有限公司","MF-0243");
		pool.put("昆山市精细化工研究所有限公司","MF-0244");
		pool.put("中国石油化工股份有限公司北京化工研究院","MF-0245");
		pool.put("Johnson Matthey Davy Technologies Ltd","MF-0246");
		pool.put("科莱恩华锦催化剂（盘锦）有限公司","MF-0247");
		pool.put("The Chemours Company FC, LLC","MF-0248");
		pool.put("Porocel Industries, LLC","MF-0249");
		pool.put("PT Clariant Adsorbents Indonesia","MF-0250");
		pool.put("朗盛化学(中国)有限公司","MF-0251");
		pool.put("ExxonMobil Catalysts and Licensing LLC","MF-0252");
		pool.put("Wacker Chemicals (China) Co., Ltd.","MF-0253");
		pool.put("禾大化学品（上海）有限公司","MF-0254");
		pool.put("Equistar Chemicals, LP","MF-0255");
		pool.put("Basell Poliolefine Italia s.r.l.","MF-0256");
		pool.put("协和化学工业株式会社","MF-0257");
		pool.put("Kyowa Chemical Industry Co,: Ltd","MF-0258");
		pool.put("Milliken Enterprise Management","MF-0259");
		pool.put("哈德逊水技术咨询有限公司","MF-0260");
		pool.put("Graver Technologies","MF-0261");
		pool.put("默克生命科学(上海)有限公司","MF-0262");
		pool.put("Aladdin Industrial Corporation","MF-0263");
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
