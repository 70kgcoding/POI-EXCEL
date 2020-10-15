package com.wdh.exceldemo.test;

        import java.util.Scanner;

public class AnswerOne {
    public static void main(String[] args) {
        Scanner in = new Scanner(System.in);
        String k = in.next();
        String n = in.next();
        String m = in.next();
        String luck = n + "";
        if (Integer.valueOf(m) <= 0 || Integer.valueOf(k) <= 0 || Integer.valueOf(n) < 0) {
            System.out.println(0);
        }
        String str = Integer.toString(Integer.valueOf(k), Integer.valueOf(m));

        System.out.println(count(str, luck));

        in.close();
    }

    public static int count(String str,String subStr) {
        int count = 0, start = 0;

        while ((start = str.indexOf(subStr,start)) >= 0) {
            start += subStr.length();
            count ++;
        }
        return count;
    }
}
