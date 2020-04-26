package com.img.utils;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * String 字符串处理 工具类
 * @author dshl
 * @date 2017-10-17
 */
public class StringUtils {

	/**
	 * 字符串转json对象
	 * @param jsonStr
	 * @return
	 */
	public static String nullToEmpty(Object object){
		if(null == object){
			return "";
		}
		return String.valueOf(object);
	}
	
	public static boolean isNullOrEmpty(String string){
		if(null == string){
			return true;
		}
		
		if(string.isEmpty()){
			return true;
		}
		
		return false;
	}

	public static String toString(Object object){
		if(null == object){
			return null;
		}
		return String.valueOf(object);
	}

	public static boolean areNotEmpty(String a, String b){
		
		return true;
	}
	
	/**
	 * 处理unicode码与其他字符串混合的字符串
	 * @param unicodeStr
	 * @return
	 */
	public static String unicodeToString(String str) {
		 
		Pattern pattern = Pattern.compile("(\\\\u(\\p{XDigit}{4}))");
		Matcher matcher = pattern.matcher(str);
		char ch;
		while (matcher.find()) {
			String group = matcher.group(2);
			ch = (char) Integer.parseInt(group, 16);
			String group1 = matcher.group(1);
			str = str.replace(group1, ch + "");
		}
		return str;

	}
	
	public static String subString(String str, String strStart, String strEnd) {
        /* 找出指定的2个字符在 该字符串里面的 位置 */
        int strStartIndex = str.indexOf(strStart);
        int strEndIndex = str.indexOf(strEnd);
 
        /* index 为负数 即表示该字符串中 没有该字符 */
        if (strStartIndex < 0) {
            return "字符串 :---->" + str + "<---- 中不存在 " + strStart + ", 无法截取目标字符串";
        }
        if (strEndIndex < 0) {
            return "字符串 :---->" + str + "<---- 中不存在 " + strEnd + ", 无法截取目标字符串";
        }
        /* 开始截取 */
        String result = str.substring(strStartIndex, strEndIndex).substring(strStart.length());
        return result;
    }

	
	public static int toInt(Object obj) {
		String str = obj != null ? obj.toString().trim() : "";
		return "".equals(str) ? 0 : Integer.parseInt(str);
	}
}