package com.linus.excel.util;

import java.sql.Date;
import java.sql.Time;
import java.text.DecimalFormat;
import java.util.logging.Logger;

public class ExcelUtil {
	private static final Logger logger = Logger.getLogger(ExcelUtil.class.getName());
	
	public static void main(String[] args) {
		Number num = 4.31423435432554E27;
		System.out.println(num);
		System.out.println(formatNumber(num));
		
		num = 88.7;
		System.out.println(num);
		System.out.println(formatNumber(num));
	}

	/**
	 * Return the mapped java class implementation of specified type. 
	 * This method support string, double, float, int, date, time, boolean, short, byte.
	 * 
	 * @param type String representation of a type.
	 * @return Java class implementation.
	 */
	public static Class<?> parseType(String type) {
		if (type == null || type.isEmpty()) return String.class;
		
		switch (type.toLowerCase()) {
		case "text"     : return String.class;
		case "double"   : return Double.class;
		case "percent"  : return Double.class;
		case "int"      : return Integer.class;
		case "date"     : return Date.class;
		case "datetime" : return Date.class;
		case "time"     : return Time.class;
		case "boolean"  : return Boolean.class;
		case "short"    : return Short.class;
		case "byte"     : return Byte.class;
		default         : return String.class;
		}
	}
	
	public static String formatNumber(Number value) {
		DecimalFormat df = new DecimalFormat("0");
		return df.format(value);
	}
}
