package com.linus.excel.util;

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
import java.util.logging.Logger;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.linus.excel.ColumnConfiguration;
import com.linus.excel.annotation.Header;
import com.linus.excel.validation.ColumnConstraint;
import com.linus.excel.validation.DoubleColumnConstraint;
import com.linus.excel.validation.IntegerRangeColumnConstraint;
import com.linus.excel.validation.LengthColumnConstraint;
import com.linus.excel.validation.NotNullColumnConstraint;
import com.linus.excel.validation.RangeColumnConstraint;
import com.linus.excel.validation.UniqueColumnConstraint;
import com.linus.locale.LocaleUtil;

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
	public static ArrayList<ColumnConfiguration> getColumnConfigurations(Class<?> clazz) throws IntrospectionException {

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
					config.setTitle(h.title());
					config.setReadOrder(h.readOrder());
					config.setWriteOrder(h.writeOrder());
					config.setWritable(h.writable());
					config.setPropertyDescriptor(descriptor);
					configs.add(config);
				}
			}
		}

		return configs;
	}
	
	/**
	 * Get column configurations form JSON array. Each element in array is a column's configuration. 
	 * But we need to convert them into ColumnConfiguration objects.
	 * 
	 * Note: It's required that each JSON object in array has following attributes: title(column title), order(column number in a row),
	 * writable(whether user can modified it), type(what data type it is), key(what name this column maps to).
	 * 
	 * @param array JSONArray
	 * @return A list of ColumnConfiguration objects.
	 */
	public static ArrayList<ColumnConfiguration> getColumnConfigurations(ArrayNode array, Locale locale) {
		ArrayList<ColumnConfiguration> configs = new ArrayList<ColumnConfiguration>();
		
		for (int i = 0; i < array.size(); i++) {
			JsonNode json = array.get(i);
			if (json != null) {
				ColumnConfiguration config = new ColumnConfiguration();
				config.setTitle(getTitle(json.get("displayLabel"), locale));
				config.setReadOrder(i);
				config.setWriteOrder(i);
				resolveCommonConfigurations(config, json);
				resolveColumnConstraint(config, json);
				configs.add(config);
			}
		}
		
		return configs;
	}
	
	/**
	 * Resolve title attribute from json node which contains column title.
	 * @param titleNode
	 * @param locale
	 * @return
	 */
	private static String getTitle(JsonNode titleNode, Locale locale) {
		if (titleNode != null && titleNode.isArray()) {
			for (int i = 0; i < titleNode.size(); i++) {
				JsonNode label = titleNode.get(i);
				if (LocaleUtil.getLocale(locale).equalsIgnoreCase(label.get("locale").asText())) {
					return label.get("labelName").asText();
				}
			}
		}
		
		return "";
	}
	
	/**
	 * Resolve title attribute from json node which contains column title.
	 * @param titleNode
	 * @param locale
	 * @return
	 */
	private static void resolveCommonConfigurations(ColumnConfiguration config, JsonNode field) {
		config.setKey(field.get("api_Name").asText());
		config.setLabel(field.get("labelName").asText());
		/*config.setDisplay(field.get("display").asBoolean());	*/
		config.setWritable(field.get("input").asBoolean());
		if (field.has("sample")) {
			String sample = field.get("sample").asText();
			config.setSample("null".equalsIgnoreCase(sample) ? "" : "Sample: " + sample);
		}
	}
	
	/**
	 * 
	 * @param config
	 * @param fieldNode
	 * @param value
	 */
	private static void resolveColumnConstraint(ColumnConfiguration config, JsonNode fieldNode) {
		boolean required = fieldNode.get("required").asBoolean();
		boolean isUnique = fieldNode.get("isUnique").asBoolean();
		
		JsonNode typeNode = fieldNode.get("fieldtype");
		if (typeNode != null && !typeNode.isNull()) {
			String type = typeNode.get("typeName").asText();
			config.setRawType(type);
			
			if (type.equalsIgnoreCase("picklist")) {
				String entries = typeNode.get("picklistEntry").asText();
				String[] list = entries.split(";");
				
				RangeColumnConstraint constraint = new RangeColumnConstraint();
				constraint.setPickList(list);
				config.getConstraints().add(constraint);
			} if (type.equalsIgnoreCase("combobox")) {
				String entries = typeNode.get("picklistEntry").asText();
				String[] list = entries.split(";");
				
				RangeColumnConstraint constraint = new RangeColumnConstraint();
				constraint.setMustInRange(false);
				constraint.setPickList(list);
				config.getConstraints().add(constraint);
			} else if (type.equalsIgnoreCase("integer")) {
				IntegerRangeColumnConstraint constraint = new IntegerRangeColumnConstraint();
				config.getConstraints().add(constraint);
			} else if (type.equalsIgnoreCase("double")) {
				DoubleColumnConstraint constraint = new DoubleColumnConstraint();
				constraint.setDigits(typeNode.get("digits").asInt(0));
				config.getConstraints().add(constraint);
			} else if (type.equalsIgnoreCase("string")) {
				config.setType(parseType(type));
				
				int length = typeNode.get("length").asInt();
				if (length > 0) {
					ColumnConstraint constraint = new LengthColumnConstraint(length);
					config.getConstraints().add(constraint);
				}
			}
			
		} else {
			// for attachment field, fieldNode is empty.
			JsonNode attachmentNode = fieldNode.get("attachmentType");
			if (attachmentNode != null && !attachmentNode.isNull()) {
				config.setRawType("attachment");
			}
		}
		
		if (required) {
			ColumnConstraint constraint = new NotNullColumnConstraint();
			config.getConstraints().add(constraint);
		}
		
		if (isUnique) {
			ColumnConstraint constraint = new UniqueColumnConstraint();
			config.getConstraints().add(constraint);
		}
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
