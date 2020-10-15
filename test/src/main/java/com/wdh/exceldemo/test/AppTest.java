package com.wdh.exceldemo.test;
import java.util.List;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedList;

public class AppTest {

	public AppTest() {
		// TODO Auto-generated constructor stub
	}
	
	public static int getDuplicateNum(int[] nums) {
		Integer[] arr = new Integer[nums.length - 1];
		for (int i = 0; i < nums.length; i++) {
			if (i == nums.length -1) {
				return nums[i];
			} else if (null == arr[nums[i]]) {
				arr[nums[i]] = nums[i];
			} else {
				return nums[i];
			}
		}
		return -1;
	}
	
    public static List<Boolean> kidsWithCandies(int[] candies, int extraCandies) {
    	List<Boolean> result = new LinkedList<Boolean>();
    	int[] arr = Arrays.copyOf(candies, candies.length);
    	Arrays.sort(arr);
    	int limit = arr[arr.length - 1] - extraCandies;
    	for (int i = 0;i < candies.length;i++) {
    		if (limit > candies[i]) {
    			result.add(false);
    		} else {
    			result.add(true);
    		}
    	}
    	return result;
    }
	
	public static void main(String[] args) {
		String str = "http://cspc-alpha.usequantum.com.cn/uploads/sds/lab-001_1,1,1-三氯乙烷.pdf";
		try {
			URL url = new URL(str);
			HttpURLConnection urlConnection = (HttpURLConnection) url.openConnection();
			urlConnection.setConnectTimeout(30000);
			urlConnection.setReadTimeout(30000);
            urlConnection.setRequestMethod("HEAD");
            urlConnection.setInstanceFollowRedirects(false);

            urlConnection.setRequestProperty("User-Agent", "Mozilla/5.0 (Windows; U; Windows NT 6.0; en-US; rv:1.9.1.2) Gecko/20090729 Firefox/3.5.2 (.NET CLR 3.5.30729)");

            String code = new Integer(urlConnection.getResponseCode()).toString();
			
			System.out.println("code -> " + code);
		} catch (MalformedURLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	
	}

}
