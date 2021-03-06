package com.wdh.exceldemo.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Hashtable;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DealWithJson2 {
	private static Hashtable<String, String> pool=new Hashtable<String, String>();
	static {
		pool.put("包头市","_SYS_A1_405_1");
		pool.put("乌海市","_SYS_A1_405_2");
		pool.put("巴彦淖尔市","_SYS_A1_405_3");
		pool.put("呼伦贝尔市","_SYS_A1_405_4");
		pool.put("鄂尔多斯市","_SYS_A1_405_5");
		pool.put("阿拉善盟","_SYS_A1_405_6");
		pool.put("赤峰市","_SYS_A1_405_7");
		pool.put("通辽市","_SYS_A1_405_8");
		pool.put("兴安盟","_SYS_A1_405_9");
		pool.put("乌兰察布市","_SYS_A1_405_10");
		pool.put("锡林郭勒盟","_SYS_A1_405_11");
		pool.put("呼和浩特市","_SYS_A1_405_12");
		pool.put("洛阳市","_SYS_A1_416_1");
		pool.put("三门峡市","_SYS_A1_416_2");
		pool.put("漯河市","_SYS_A1_416_3");
		pool.put("许昌市","_SYS_A1_416_4");
		pool.put("南阳市","_SYS_A1_416_5");
		pool.put("信阳市","_SYS_A1_416_6");
		pool.put("济源市","_SYS_A1_416_7");
		pool.put("驻马店市","_SYS_A1_416_8");
		pool.put("濮阳市","_SYS_A1_416_9");
		pool.put("焦作市","_SYS_A1_416_10");
		pool.put("鹤壁市","_SYS_A1_416_11");
		pool.put("新乡市","_SYS_A1_416_12");
		pool.put("平顶山市","_SYS_A1_416_13");
		pool.put("周口市","_SYS_A1_416_14");
		pool.put("商丘市","_SYS_A1_416_15");
		pool.put("开封市","_SYS_A1_416_16");
		pool.put("郑州市","_SYS_A1_416_17");
		pool.put("安阳市","_SYS_A1_416_18");
		pool.put("北屯市","_SYS_A1_431_1");
		pool.put("铁门关市","_SYS_A1_431_2");
		pool.put("双河市","_SYS_A1_431_3");
		pool.put("可克达拉市","_SYS_A1_431_4");
		pool.put("博尔塔拉蒙古自治州","_SYS_A1_431_5");
		pool.put("塔城地区","_SYS_A1_431_6");
		pool.put("和田地区","_SYS_A1_431_7");
		pool.put("阿勒泰地区","_SYS_A1_431_8");
		pool.put("昆玉市","_SYS_A1_431_9");
		pool.put("克拉玛依市","_SYS_A1_431_10");
		pool.put("石河子市","_SYS_A1_431_11");
		pool.put("昌吉回族自治州","_SYS_A1_431_12");
		pool.put("五家渠市","_SYS_A1_431_13");
		pool.put("巴音郭楞蒙古自治州","_SYS_A1_431_14");
		pool.put("乌鲁木齐市","_SYS_A1_431_15");
		pool.put("伊犁哈萨克自治州","_SYS_A1_431_16");
		pool.put("阿克苏地区","_SYS_A1_431_17");
		pool.put("阿拉尔市","_SYS_A1_431_18");
		pool.put("图木舒克市","_SYS_A1_431_19");
		pool.put("喀什地区","_SYS_A1_431_20");
		pool.put("克孜勒苏柯尔克孜自治州","_SYS_A1_431_21");
		pool.put("哈密市","_SYS_A1_431_22");
		pool.put("吐鲁番市","_SYS_A1_431_23");
		pool.put("葫芦岛市","_SYS_A1_406_1");
		pool.put("大连市","_SYS_A1_406_2");
		pool.put("丹东市","_SYS_A1_406_3");
		pool.put("锦州市","_SYS_A1_406_4");
		pool.put("抚顺市","_SYS_A1_406_5");
		pool.put("沈阳市","_SYS_A1_406_6");
		pool.put("铁岭市","_SYS_A1_406_7");
		pool.put("营口市","_SYS_A1_406_8");
		pool.put("朝阳市","_SYS_A1_406_9");
		pool.put("辽阳市","_SYS_A1_406_10");
		pool.put("鞍山市","_SYS_A1_406_11");
		pool.put("阜新市","_SYS_A1_406_12");
		pool.put("盘锦市","_SYS_A1_406_13");
		pool.put("本溪市","_SYS_A1_406_14");
		pool.put("十堰市","_SYS_A1_417_1");
		pool.put("襄阳市","_SYS_A1_417_2");
		pool.put("荆门市","_SYS_A1_417_3");
		pool.put("宜昌市","_SYS_A1_417_4");
		pool.put("武汉市","_SYS_A1_417_5");
		pool.put("黄冈市","_SYS_A1_417_6");
		pool.put("天门市","_SYS_A1_417_7");
		pool.put("孝感市","_SYS_A1_417_8");
		pool.put("潜江市","_SYS_A1_417_9");
		pool.put("恩施土家族苗族自治州","_SYS_A1_417_10");
		pool.put("仙桃市","_SYS_A1_417_11");
		pool.put("荆州市","_SYS_A1_417_12");
		pool.put("咸宁市","_SYS_A1_417_13");
		pool.put("神农架林区","_SYS_A1_417_14");
		pool.put("随州市","_SYS_A1_417_15");
		pool.put("鄂州市","_SYS_A1_417_16");
		pool.put("黄石市","_SYS_A1_417_17");
		pool.put("汕头市","_SYS_A1_419_1");
		pool.put("佛山市","_SYS_A1_419_2");
		pool.put("肇庆市","_SYS_A1_419_3");
		pool.put("惠州市","_SYS_A1_419_4");
		pool.put("深圳市","_SYS_A1_419_5");
		pool.put("珠海市","_SYS_A1_419_6");
		pool.put("湛江市","_SYS_A1_419_7");
		pool.put("江门市","_SYS_A1_419_8");
		pool.put("阳江市","_SYS_A1_419_9");
		pool.put("茂名市","_SYS_A1_419_10");
		pool.put("东沙群岛","_SYS_A1_419_11");
		pool.put("汕尾市","_SYS_A1_419_12");
		pool.put("潮州市","_SYS_A1_419_13");
		pool.put("云浮市","_SYS_A1_419_14");
		pool.put("河源市","_SYS_A1_419_15");
		pool.put("梅州市","_SYS_A1_419_16");
		pool.put("东莞市","_SYS_A1_419_17");
		pool.put("清远市","_SYS_A1_419_18");
		pool.put("广州市","_SYS_A1_419_19");
		pool.put("揭阳市","_SYS_A1_419_20");
		pool.put("韶关市","_SYS_A1_419_21");
		pool.put("中山市","_SYS_A1_419_22");
		pool.put("大兴安岭地区","_SYS_A1_408_1");
		pool.put("七台河市","_SYS_A1_408_2");
		pool.put("鹤岗市","_SYS_A1_408_3");
		pool.put("伊春市","_SYS_A1_408_4");
		pool.put("绥化市","_SYS_A1_408_5");
		pool.put("哈尔滨市","_SYS_A1_408_6");
		pool.put("黑河市","_SYS_A1_408_7");
		pool.put("齐齐哈尔市","_SYS_A1_408_8");
		pool.put("牡丹江市","_SYS_A1_408_9");
		pool.put("鸡西市","_SYS_A1_408_10");
		pool.put("大庆市","_SYS_A1_408_11");
		pool.put("双鸭山市","_SYS_A1_408_12");
		pool.put("佳木斯市","_SYS_A1_408_13");
		pool.put("连云港市","_SYS_A1_410_1");
		pool.put("宿迁市","_SYS_A1_410_2");
		pool.put("南京市","_SYS_A1_410_3");
		pool.put("镇江市","_SYS_A1_410_4");
		pool.put("南通市","_SYS_A1_410_5");
		pool.put("淮安市","_SYS_A1_410_6");
		pool.put("徐州市","_SYS_A1_410_7");
		pool.put("盐城市","_SYS_A1_410_8");
		pool.put("泰州市","_SYS_A1_410_9");
		pool.put("扬州市","_SYS_A1_410_10");
		pool.put("无锡市","_SYS_A1_410_11");
		pool.put("常州市","_SYS_A1_410_12");
		pool.put("苏州市","_SYS_A1_410_13");
		pool.put("商洛市","_SYS_A1_427_1");
		pool.put("西安市","_SYS_A1_427_2");
		pool.put("汉中市","_SYS_A1_427_3");
		pool.put("铜川市","_SYS_A1_427_4");
		pool.put("安康市","_SYS_A1_427_5");
		pool.put("榆林市","_SYS_A1_427_6");
		pool.put("延安市","_SYS_A1_427_7");
		pool.put("宝鸡市","_SYS_A1_427_8");
		pool.put("咸阳市","_SYS_A1_427_9");
		pool.put("渭南市","_SYS_A1_427_10");
		pool.put("烟台市","_SYS_A1_415_1");
		pool.put("威海市","_SYS_A1_415_2");
		pool.put("青岛市","_SYS_A1_415_3");
		pool.put("淄博市","_SYS_A1_415_4");
		pool.put("聊城市","_SYS_A1_415_5");
		pool.put("临沂市","_SYS_A1_415_6");
		pool.put("潍坊市","_SYS_A1_415_7");
		pool.put("东营市","_SYS_A1_415_8");
		pool.put("滨州市","_SYS_A1_415_9");
		pool.put("枣庄市","_SYS_A1_415_10");
		pool.put("日照市","_SYS_A1_415_11");
		pool.put("济宁市","_SYS_A1_415_12");
		pool.put("泰安市","_SYS_A1_415_13");
		pool.put("德州市","_SYS_A1_415_14");
		pool.put("济南市","_SYS_A1_415_15");
		pool.put("菏泽市","_SYS_A1_415_16");
		pool.put("昌都市","_SYS_A1_426_1");
		pool.put("那曲市","_SYS_A1_426_2");
		pool.put("拉萨市","_SYS_A1_426_3");
		pool.put("日喀则市","_SYS_A1_426_4");
		pool.put("山南市","_SYS_A1_426_5");
		pool.put("林芝市","_SYS_A1_426_6");
		pool.put("阿里地区","_SYS_A1_426_7");
		pool.put("宁德市","_SYS_A1_413_1");
		pool.put("福州市","_SYS_A1_413_2");
		pool.put("龙岩市","_SYS_A1_413_3");
		pool.put("莆田市","_SYS_A1_413_4");
		pool.put("泉州市","_SYS_A1_413_5");
		pool.put("厦门市","_SYS_A1_413_6");
		pool.put("漳州市","_SYS_A1_413_7");
		pool.put("南平市","_SYS_A1_413_8");
		pool.put("三明市","_SYS_A1_413_9");
		pool.put("淮北市","_SYS_A1_412_1");
		pool.put("阜阳市","_SYS_A1_412_2");
		pool.put("马鞍山市","_SYS_A1_412_3");
		pool.put("铜陵市","_SYS_A1_412_4");
		pool.put("池州市","_SYS_A1_412_5");
		pool.put("亳州市","_SYS_A1_412_6");
		pool.put("蚌埠市","_SYS_A1_412_7");
		pool.put("滁州市","_SYS_A1_412_8");
		pool.put("六安市","_SYS_A1_412_9");
		pool.put("安庆市","_SYS_A1_412_10");
		pool.put("黄山市","_SYS_A1_412_11");
		pool.put("宣城市","_SYS_A1_412_12");
		pool.put("淮南市","_SYS_A1_412_13");
		pool.put("合肥市","_SYS_A1_412_14");
		pool.put("宿州市","_SYS_A1_412_15");
		pool.put("芜湖市","_SYS_A1_412_16");
		pool.put("遵义市","_SYS_A1_424_1");
		pool.put("铜仁市","_SYS_A1_424_2");
		pool.put("六盘水市","_SYS_A1_424_3");
		pool.put("黔东南苗族侗族自治州","_SYS_A1_424_4");
		pool.put("黔南布依族苗族自治州","_SYS_A1_424_5");
		pool.put("安顺市","_SYS_A1_424_6");
		pool.put("黔西南布依族苗族自治州","_SYS_A1_424_7");
		pool.put("毕节市","_SYS_A1_424_8");
		pool.put("贵阳市","_SYS_A1_424_9");
		pool.put("重庆郊县","_SYS_A1_422_1");
		pool.put("重庆城区","_SYS_A1_422_2");
		pool.put("上海城区","_SYS_A1_409_1");
		pool.put("岳阳市","_SYS_A1_418_1");
		pool.put("益阳市","_SYS_A1_418_2");
		pool.put("衡阳市","_SYS_A1_418_3");
		pool.put("娄底市","_SYS_A1_418_4");
		pool.put("长沙市","_SYS_A1_418_5");
		pool.put("张家界市","_SYS_A1_418_6");
		pool.put("怀化市","_SYS_A1_418_7");
		pool.put("湘西土家族苗族自治州","_SYS_A1_418_8");
		pool.put("常德市","_SYS_A1_418_9");
		pool.put("株洲市","_SYS_A1_418_10");
		pool.put("邵阳市","_SYS_A1_418_11");
		pool.put("湘潭市","_SYS_A1_418_12");
		pool.put("永州市","_SYS_A1_418_13");
		pool.put("郴州市","_SYS_A1_418_14");
		pool.put("临高县","_SYS_A1_421_1");
		pool.put("定安县","_SYS_A1_421_2");
		pool.put("屯昌县","_SYS_A1_421_3");
		pool.put("昌江黎族自治县","_SYS_A1_421_4");
		pool.put("白沙黎族自治县","_SYS_A1_421_5");
		pool.put("琼海市","_SYS_A1_421_6");
		pool.put("琼中黎族苗族自治县","_SYS_A1_421_7");
		pool.put("东方市","_SYS_A1_421_8");
		pool.put("万宁市","_SYS_A1_421_9");
		pool.put("五指山市","_SYS_A1_421_10");
		pool.put("乐东黎族自治县","_SYS_A1_421_11");
		pool.put("保亭黎族苗族自治县","_SYS_A1_421_12");
		pool.put("陵水黎族自治县","_SYS_A1_421_13");
		pool.put("三沙市","_SYS_A1_421_14");
		pool.put("文昌市","_SYS_A1_421_15");
		pool.put("儋州市","_SYS_A1_421_16");
		pool.put("澄迈县","_SYS_A1_421_17");
		pool.put("三亚市","_SYS_A1_421_18");
		pool.put("海口市","_SYS_A1_421_19");
		pool.put("百色市","_SYS_A1_420_1");
		pool.put("钦州市","_SYS_A1_420_2");
		pool.put("北海市","_SYS_A1_420_3");
		pool.put("桂林市","_SYS_A1_420_4");
		pool.put("河池市","_SYS_A1_420_5");
		pool.put("柳州市","_SYS_A1_420_6");
		pool.put("来宾市","_SYS_A1_420_7");
		pool.put("南宁市","_SYS_A1_420_8");
		pool.put("崇左市","_SYS_A1_420_9");
		pool.put("防城港市","_SYS_A1_420_10");
		pool.put("贺州市","_SYS_A1_420_11");
		pool.put("玉林市","_SYS_A1_420_12");
		pool.put("贵港市","_SYS_A1_420_13");
		pool.put("梧州市","_SYS_A1_420_14");
		pool.put("唐山市","_SYS_A1_403_1");
		pool.put("承德市","_SYS_A1_403_2");
		pool.put("廊坊市","_SYS_A1_403_3");
		pool.put("秦皇岛市","_SYS_A1_403_4");
		pool.put("保定市","_SYS_A1_403_5");
		pool.put("石家庄市","_SYS_A1_403_6");
		pool.put("邢台市","_SYS_A1_403_7");
		pool.put("邯郸市","_SYS_A1_403_8");
		pool.put("张家口市","_SYS_A1_403_9");
		pool.put("沧州市","_SYS_A1_403_10");
		pool.put("衡水市","_SYS_A1_403_11");
		pool.put("海东市","_SYS_A1_429_1");
		pool.put("海南藏族自治州","_SYS_A1_429_2");
		pool.put("海西蒙古族藏族自治州","_SYS_A1_429_3");
		pool.put("玉树藏族自治州","_SYS_A1_429_4");
		pool.put("黄南藏族自治州","_SYS_A1_429_5");
		pool.put("果洛藏族自治州","_SYS_A1_429_6");
		pool.put("西宁市","_SYS_A1_429_7");
		pool.put("海北藏族自治州","_SYS_A1_429_8");
		pool.put("北区","_SYS_A1_433_1");
		pool.put("大埔区","_SYS_A1_433_2");
		pool.put("西贡区","_SYS_A1_433_3");
		pool.put("沙田区","_SYS_A1_433_4");
		pool.put("屯门区","_SYS_A1_433_5");
		pool.put("黄大仙区","_SYS_A1_433_6");
		pool.put("九龙城区","_SYS_A1_433_7");
		pool.put("深水埗区","_SYS_A1_433_8");
		pool.put("观塘区","_SYS_A1_433_9");
		pool.put("油尖旺区","_SYS_A1_433_10");
		pool.put("离岛区","_SYS_A1_433_11");
		pool.put("东区","_SYS_A1_433_12");
		pool.put("中西区","_SYS_A1_433_13");
		pool.put("湾仔区","_SYS_A1_433_14");
		pool.put("南区","_SYS_A1_433_15");
		pool.put("元朗区","_SYS_A1_433_16");
		pool.put("荃湾区","_SYS_A1_433_17");
		pool.put("葵青区","_SYS_A1_433_18");
		pool.put("舟山市","_SYS_A1_411_1");
		pool.put("嘉兴市","_SYS_A1_411_2");
		pool.put("宁波市","_SYS_A1_411_3");
		pool.put("台州市","_SYS_A1_411_4");
		pool.put("温州市","_SYS_A1_411_5");
		pool.put("丽水市","_SYS_A1_411_6");
		pool.put("杭州市","_SYS_A1_411_7");
		pool.put("绍兴市","_SYS_A1_411_8");
		pool.put("湖州市","_SYS_A1_411_9");
		pool.put("衢州市","_SYS_A1_411_10");
		pool.put("金华市","_SYS_A1_411_11");
		pool.put("九江市","_SYS_A1_414_1");
		pool.put("新余市","_SYS_A1_414_2");
		pool.put("抚州市","_SYS_A1_414_3");
		pool.put("鹰潭市","_SYS_A1_414_4");
		pool.put("赣州市","_SYS_A1_414_5");
		pool.put("南昌市","_SYS_A1_414_6");
		pool.put("宜春市","_SYS_A1_414_7");
		pool.put("萍乡市","_SYS_A1_414_8");
		pool.put("吉安市","_SYS_A1_414_9");
		pool.put("景德镇市","_SYS_A1_414_10");
		pool.put("上饶市","_SYS_A1_414_11");
		pool.put("嘉峪关市","_SYS_A1_428_1");
		pool.put("酒泉市","_SYS_A1_428_2");
		pool.put("金昌市","_SYS_A1_428_3");
		pool.put("兰州市","_SYS_A1_428_4");
		pool.put("平凉市","_SYS_A1_428_5");
		pool.put("白银市","_SYS_A1_428_6");
		pool.put("天水市","_SYS_A1_428_7");
		pool.put("武威市","_SYS_A1_428_8");
		pool.put("陇南市","_SYS_A1_428_9");
		pool.put("甘南藏族自治州","_SYS_A1_428_10");
		pool.put("临夏回族自治州","_SYS_A1_428_11");
		pool.put("张掖市","_SYS_A1_428_12");
		pool.put("庆阳市","_SYS_A1_428_13");
		pool.put("定西市","_SYS_A1_428_14");
		pool.put("风顺堂区","_SYS_A1_434_1");
		pool.put("路凼填海区","_SYS_A1_434_2");
		pool.put("望德堂区","_SYS_A1_434_3");
		pool.put("嘉模堂区","_SYS_A1_434_4");
		pool.put("花地玛堂区","_SYS_A1_434_5");
		pool.put("花王堂区","_SYS_A1_434_6");
		pool.put("圣方济各堂区","_SYS_A1_434_7");
		pool.put("大堂区","_SYS_A1_434_8");
		pool.put("固原市","_SYS_A1_430_1");
		pool.put("中卫市","_SYS_A1_430_2");
		pool.put("银川市","_SYS_A1_430_3");
		pool.put("石嘴山市","_SYS_A1_430_4");
		pool.put("吴忠市","_SYS_A1_430_5");
		pool.put("广元市","_SYS_A1_423_1");
		pool.put("南充市","_SYS_A1_423_2");
		pool.put("巴中市","_SYS_A1_423_3");
		pool.put("德阳市","_SYS_A1_423_4");
		pool.put("绵阳市","_SYS_A1_423_5");
		pool.put("成都市","_SYS_A1_423_6");
		pool.put("广安市","_SYS_A1_423_7");
		pool.put("达州市","_SYS_A1_423_8");
		pool.put("遂宁市","_SYS_A1_423_9");
		pool.put("资阳市","_SYS_A1_423_10");
		pool.put("眉山市","_SYS_A1_423_11");
		pool.put("内江市","_SYS_A1_423_12");
		pool.put("乐山市","_SYS_A1_423_13");
		pool.put("自贡市","_SYS_A1_423_14");
		pool.put("泸州市","_SYS_A1_423_15");
		pool.put("宜宾市","_SYS_A1_423_16");
		pool.put("凉山彝族自治州","_SYS_A1_423_17");
		pool.put("攀枝花市","_SYS_A1_423_18");
		pool.put("阿坝藏族羌族自治州","_SYS_A1_423_19");
		pool.put("雅安市","_SYS_A1_423_20");
		pool.put("甘孜藏族自治州","_SYS_A1_423_21");
		pool.put("吉林市","_SYS_A1_407_1");
		pool.put("长春市","_SYS_A1_407_2");
		pool.put("白城市","_SYS_A1_407_3");
		pool.put("松原市","_SYS_A1_407_4");
		pool.put("辽源市","_SYS_A1_407_5");
		pool.put("四平市","_SYS_A1_407_6");
		pool.put("延边朝鲜族自治州","_SYS_A1_407_7");
		pool.put("白山市","_SYS_A1_407_8");
		pool.put("通化市","_SYS_A1_407_9");
		pool.put("昭通市","_SYS_A1_425_1");
		pool.put("曲靖市","_SYS_A1_425_2");
		pool.put("红河哈尼族彝族自治州","_SYS_A1_425_3");
		pool.put("怒江傈僳族自治州","_SYS_A1_425_4");
		pool.put("西双版纳傣族自治州","_SYS_A1_425_5");
		pool.put("玉溪市","_SYS_A1_425_6");
		pool.put("大理白族自治州","_SYS_A1_425_7");
		pool.put("丽江市","_SYS_A1_425_8");
		pool.put("迪庆藏族自治州","_SYS_A1_425_9");
		pool.put("文山壮族苗族自治州","_SYS_A1_425_10");
		pool.put("保山市","_SYS_A1_425_11");
		pool.put("普洱市","_SYS_A1_425_12");
		pool.put("昆明市","_SYS_A1_425_13");
		pool.put("楚雄彝族自治州","_SYS_A1_425_14");
		pool.put("临沧市","_SYS_A1_425_15");
		pool.put("德宏傣族景颇族自治州","_SYS_A1_425_16");
		pool.put("阳泉市","_SYS_A1_404_1");
		pool.put("太原市","_SYS_A1_404_2");
		pool.put("临汾市","_SYS_A1_404_3");
		pool.put("运城市","_SYS_A1_404_4");
		pool.put("长治市","_SYS_A1_404_5");
		pool.put("朔州市","_SYS_A1_404_6");
		pool.put("忻州市","_SYS_A1_404_7");
		pool.put("晋城市","_SYS_A1_404_8");
		pool.put("晋中市","_SYS_A1_404_9");
		pool.put("吕梁市","_SYS_A1_404_10");
		pool.put("大同市","_SYS_A1_404_11");
		pool.put("北京城区","_SYS_A1_401_1");
		pool.put("天津城区","_SYS_A1_402_1");
	}
	public static void excelRead() throws Exception {
		//获得Excel文件输出流

		InputStream in = new FileInputStream("D:\\tmp/testforjson.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //遍历行
    	for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
    		XSSFRow row = sheet.getRow(i);
    		if (row.getCell(2)!= null && row.getCell(2).getStringCellValue() != null) {
    			row.getCell(2).setCellValue(pool.get(row.getCell(2).getStringCellValue()));
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
