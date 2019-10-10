package com.linus.excel.po.util;

import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Method;
import java.sql.Date;
import java.sql.Time;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Locale;
import java.util.ResourceBundle;
import java.util.logging.Logger;

import com.linus.excel.ColumnConfiguration;
import com.linus.excel.annotation.Header;

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
	 * Get column configurations for POJO properties. These configurations are stored in <code>Header</code> annotation.
	 * 
	 * @see Header
	 * 
	 * @param clazz Class
	 * @return A list of column configuration on properties.
	 * @throws IntrospectionException
	 */
	public static ArrayList<ColumnConfiguration> getColumnConfigurations(Class<?> clazz, Locale locale, ResourceBundle bundle) throws IntrospectionException {

		BeanInfo info = Introspector.getBeanInfo(clazz);
		PropertyDescriptor[] descriptors = info.getPropertyDescriptors();
		ArrayList<ColumnConfiguration> configs = new ArrayList<ColumnConfiguration>();

		for (int i = 0; i < descriptors.length; i++) {
			PropertyDescriptor descriptor = descriptors[i];
			Method getter = descriptor.getReadMethod();
			// "Boolean" type property doesn't support "is" getter. Only "boolean" supports. 
			if (getter != null) {
				Header h = descriptor.getReadMethod().getAnnotation(Header.class);
				if (h != null) {
					ColumnConfiguration config = new ColumnConfiguration();
					config.setTitle(bundle.getString(h.title()));
					config.setKey(descriptor.getName());
					config.setColumnIndex(h.columnIndex());
					config.setWritable(h.writable());
					config.setRawType(h.rawType());
					config.setPropertyDescriptor(descriptor);
					configs.add(config);
				}
			}
		}

		return configs;
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
