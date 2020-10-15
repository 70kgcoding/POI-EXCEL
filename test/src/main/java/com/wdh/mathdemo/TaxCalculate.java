package com.wdh.mathdemo;



import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

public class TaxCalculate {
	static Map<String, Integer> map4Tax = new HashMap<String, Integer>();
	static {
		map4Tax.put("污染物1", 1);
		map4Tax.put("污染物2", 2);
		map4Tax.put("污染物3", 3);
		map4Tax.put("污染物4", 4);
		map4Tax.put("污染物5", 5);
		map4Tax.put("污染物6", 6);
	}
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Map<String, Integer> map = new HashMap<String, Integer>();
		map.put("污染物1", 12);
		map.put("污染物2", 8);
		map.put("污染物3", 18);
		map.put("污染物4", 11);
		map.put("污染物5", 15);
		map.put("污染物6", 6);
		List<Entry<String, Integer>> list = new ArrayList<Entry<String, Integer>>(map.entrySet());
		Collections.sort(list, new Comparator<Map.Entry<String, Integer>>() {
			public int compare(Map.Entry<String, Integer> o1, Map.Entry<String, Integer> o2) {
				return (o2.getValue() - o1.getValue());
			}
		});
		int sum = 0;
		for (int i=0;i<3;i++) {
			System.out.println(list.get(i).getKey()+ " " +list.get(i).getValue());
			sum += list.get(i).getValue()*map4Tax.get(list.get(i).getKey());
		}
		System.out.println("税" +sum);
	}

}

 
