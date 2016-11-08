package com.linus.excel;

import java.beans.IntrospectionException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.sql.Time;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.validation.ConstraintViolation;
import javax.validation.Validator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;

import com.linus.date.DateUtil;
import com.linus.enums.ICustomEnum;
import com.linus.excel.util.ExcelUtil;
import com.linus.excel.validation.ExcelValidator;

/**
 * SheetReader.readSheet() support localization messages, default message bundle is "ValidationMessages.properties".
 * 
 * @author lyan2
 */
public class SheetReader implements ISheetReader {

	private final Logger logger = Logger.getLogger(SheetReader.class.getName());
	protected SimpleDateFormat timeformat = new SimpleDateFormat("HH:mm:ss");
	public static Map<Type, ArrayList<ColumnConfiguration>> sheetHeaders = new ConcurrentHashMap<Type, ArrayList<ColumnConfiguration>>();
	
	private Validator validator;
	
	/**
	 * Remove scientific notation
	 * @param value
	 * @return
	 */
	public static String formatNumber(Number value) {
		DecimalFormat df = new DecimalFormat("0");
		return df.format(value);
	}
	
	@Override
	public List<List<Object>> readSheet(Sheet sheet, int firstRowNum) {
		if (sheet == null) return null;
		
		ArrayList<List<Object>> list = new ArrayList<List<Object>>();
		int lastRowNum = sheet.getLastRowNum();
		
		for (int i = firstRowNum; i <= lastRowNum; i++) {
			logger.log(Level.INFO,	"Beginning to read row " + i);
			
			List<Object> obj = readRow(sheet.getRow(i));
			if (obj != null) {
				list.add(obj);
			}
		}
		
		return list;
	}

	
	@Override
	public List<List<Object>> readSheet(Sheet sheet, int firstRowNum, short firstCellNum, short lastCellNum) {
		
		return readSheet(sheet, firstRowNum, sheet.getLastRowNum(), firstCellNum, lastCellNum);
	}

	@Override
	public List<List<Object>> readSheet(Sheet sheet, int firstRowNum, int lastRowNum, short firstCellNum, short lastCellNum) {
		
		if (sheet == null) return null;
		if (lastRowNum < firstRowNum) return null;
		if (lastCellNum < firstCellNum) return null;
		
		ArrayList<List<Object>> list = new ArrayList<List<Object>>();
		
		for (int i = firstRowNum; i <= lastRowNum; i++) {
			logger.log(Level.INFO,	"Beginning to read row " + i);
			
			List<Object> rowData = readRow(sheet.getRow(i), firstCellNum, lastCellNum);
			if (rowData != null) {
				list.add(rowData);
			}
		}
		
		return list;
	}
	
	@Override
	public List<Map<String, Object>> readSheet(Sheet sheet,
			List<ColumnConfiguration> headers, int firstRowNum,
			Set<ConstraintViolation<Object>> violations) {
		if (sheet == null) return null;
		
		return readSheet(sheet, headers, firstRowNum, sheet.getLastRowNum(), violations);
	}

	@Override
	public List<Map<String, Object>> readSheet(Sheet sheet,
			List<ColumnConfiguration> configs, int firstRowNum, int lastRowNum,
			Set<ConstraintViolation<Object>> violations) {
		if (sheet == null) return null;
		
		ArrayList<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		ExcelValidator validator = new ExcelValidator();
		
		for (int i = firstRowNum; i < lastRowNum; i++) {
			logger.log(Level.INFO,	"Beginning to read row " + i);
			
			Map<String, Object> obj = readRow(configs, sheet.getRow(i));
			if (obj != null) {
				violations.addAll(validator.validate(i, obj, configs));
				if (violations.size() <= 0) {
					list.add(obj);
				} else {
					break;
				}
			}
		}
		
		return list;
	}
	
	@Override
	public List<List<Object>> readSheet2 (Sheet sheet,
			List<ColumnConfiguration> configs, int firstRowNum, Set<ConstraintViolation<Object>> violations) {
		if (sheet == null) return null;
		
		ArrayList<List<Object>> list = new ArrayList<List<Object>>();
		ExcelValidator validator = new ExcelValidator();
		int lastRowNum = sheet.getLastRowNum();
		
		for (int i = firstRowNum; i < lastRowNum; i++) {
			logger.log(Level.INFO,	"Beginning to read row " + i);
			
			List<Object> obj = readRow(sheet.getRow(i));
			if (obj != null) {
				violations.addAll(validator.validate(i, obj, configs));
				if (violations.size() <= 0) {
					list.add(obj);
				} else {
					break;
				}				
			}
		}
		
		return list;
	}
	
	@Override
	public List<List<Object>> readSheet2 (Sheet sheet,
			List<ColumnConfiguration> configs, int firstRowNum, int lastRowNum,
			Set<ConstraintViolation<Object>> violations) {
		if (sheet == null) return null;
		
		ArrayList<List<Object>> list = new ArrayList<List<Object>>();
		ExcelValidator validator = new ExcelValidator();
		
		for (int i = firstRowNum; i < lastRowNum; i++) {
			logger.log(Level.INFO,	"Beginning to read row " + i);
			
			List<Object> obj = readRow(sheet.getRow(i));
			if (obj != null) {
				violations.addAll(validator.validate(i, obj, configs));
				if (violations.size() <= 0) {
					list.add(obj);
				} else {
					break;
				}				
			}
		}
		
		return list;
	}

	
	@Override
	public List<Object> readSheet(Sheet sheet, Class<?> clazz, int firstDataRow, Set<ConstraintViolation<Object>> constraintViolations) {
				
		ArrayList<ColumnConfiguration> headers = null;
		ArrayList<Object> list = new ArrayList<Object>();

		if (sheetHeaders.containsKey(clazz)) {
			 headers = sheetHeaders.get(clazz);
		} else {
			try {
				headers = ExcelUtil.getColumnConfigurations(clazz);
			} catch (IntrospectionException e) {
				logger.log(Level.SEVERE, "Failed to get column configuration from class - " + clazz.getName() + ", due to bean instropection exception.");
			}
			
			sheetHeaders.put(clazz, headers);
		}

		if (headers != null) {
			
			int lastRowNum = sheet.getLastRowNum();
			
			for (int i = firstDataRow; i < lastRowNum; i++) {
				logger.log(Level.INFO,	"Beginning to read row " + i);
				Row row = sheet.getRow(i);
				
				if (row != null) {
					Object obj = readRow(headers, row, clazz);
					if (obj != null) {
						if (validator != null) {
							Set<ConstraintViolation<Object>> violations = validator.validate(obj);
							
							if (violations == null || violations.isEmpty()) {
								list.add(obj);
							} else {
								if (constraintViolations != null) {
									constraintViolations.addAll(violations);
									break;
								}
							}
						} else {
							list.add(obj);
						}
					}
				}
			}
			
			return list;

		}

		return null;
	}
	


	
	@Override
	public List<Object> readRow(Row row) {
		if (row == null) return null;
		
		return readRow(row, (short)0, row.getLastCellNum());
	}

	@Override
	public List<Object> readRow(Row row, short firstCellNum) {
		if (row == null) return null;
		
		return readRow(row, firstCellNum, row.getLastCellNum());
	}

	@Override
	public List<Object> readRow(Row row, short firstCellNum, short lastCellNum) {
		if (row == null) return null;
		
		boolean isAllNull = true;
		
		List<Object> list = new ArrayList<Object>();
		for (short i = firstCellNum; i < lastCellNum; i++) {
			/** A new, blank cell is created for missing cells. Blank cells are returned as normal */
			Object value = readCell(row.getCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK));
			list.add(value);
			
			if (value != null) isAllNull = false;
		}
		
		return isAllNull ? null : list;
	}

	@Override
	public Map<String, Object> readRow(List<ColumnConfiguration> headers, Row row) {
		if (row == null) return null;
		
		Map<String, Object> map = new HashMap<String, Object>();
		Iterator<ColumnConfiguration> iter = headers.iterator();
		
		while (iter.hasNext()) {
			ColumnConfiguration header = iter.next();
			
			// if cell doesn't exist, return null;
			Cell cell = row.getCell(header.getReadOrder(), MissingCellPolicy.CREATE_NULL_AS_BLANK);
			Object value = null;
			
			if (header.getType() != null) {
				value = readCell(cell, header.getType());
			} else {
				if (!"attachment".equalsIgnoreCase(header.getRawType())) {
					value = readCell(cell);
				}
			}		
			
			map.put(header.getKey(), value);
		}
		
		return map;
	}

	
	/**
	 * Read a single row and convert it into specified type instance. All cell values will be instance properties. 
	 */
	public Object readRow(List<ColumnConfiguration> headers, Row row,
			Class<?> clazz) {
		
		Object o = null;

		if (headers != null && !headers.isEmpty()) {
			try {
				o = clazz.newInstance();
			} catch (InstantiationException e) {
				logger.log(Level.WARNING, "Class " + clazz.getName()
						+ " can't be instantiated!");
				e.printStackTrace();
				return o;
			} catch (IllegalAccessException e) {
				logger.log(Level.WARNING, "Class " + clazz.getName()
						+ " can't be instantiated! Because it's not accssable.");
				e.printStackTrace();
				return o;
			}

			Iterator<ColumnConfiguration> iter = headers.iterator();
			ColumnConfiguration header = null;
			while (iter.hasNext()) {
				header = iter.next();
				
				Cell cell = row.getCell(header.getReadOrder(), MissingCellPolicy.RETURN_NULL_AND_BLANK);
				Method setter = header.getPropertyDescriptor().getWriteMethod();
				Object value = readCell(cell, header.getPropertyDescriptor()
						.getReadMethod().getReturnType());

				logger.log(Level.INFO, "Property "
						+ header.getPropertyDescriptor().getName()
						+ " get value " + value);
				try {
					if (value != null ) {
						setter.invoke(o, value);
					}
				} catch (IllegalAccessException | IllegalArgumentException
						| InvocationTargetException e) {
					logger.log(Level.WARNING, "Property "
							+ header.getPropertyDescriptor().getName()
							+ " can't be set on Class " + clazz.getName()
							+ " instance, value is " + value);
					e.printStackTrace();
				}
			}
		}
		return o;
	}
	
	@Override
	public Object readCell(Cell cell) {
		if (cell == null) return null;
		
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_BLANK: return null;
		case Cell.CELL_TYPE_ERROR: return null;
		case Cell.CELL_TYPE_NUMERIC:
			return cell.getNumericCellValue();
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_BOOLEAN: 
			return cell.getBooleanCellValue();
		case Cell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();
		default: 
			return null;
		}
	}
	
	/**
	 * Read cell value and convert them into specified type.
	 * @param cell
	 * @param type
	 * @return
	 */
	public Object readCell(Cell cell, Class<?> type) {
		if (cell == null) return null;
		
		logger.log(Level.INFO, "Cell type " + cell.getCellType()
				+ ", property type " + type);
		
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_BLANK: return null;
		case Cell.CELL_TYPE_ERROR: return null;
		case Cell.CELL_TYPE_NUMERIC:
			return readFromNumberCell(cell, type);
		case Cell.CELL_TYPE_STRING:
			return readFromStringCell(cell, type);
		case Cell.CELL_TYPE_BOOLEAN: 
			Boolean value = cell.getBooleanCellValue();
			if (Boolean.class.isAssignableFrom(type)) {
				return value;
			} else {
				return value.toString();
			}
		case Cell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();
		default: 
			return null;
		}
	}
	
	/**
	 * If cell type is number, its value is Double type.
	 * @param cellVal
	 * @param type
	 * @return
	 */
	protected Object resolveDouble(Double cellVal, Class<?> type) {
		if (Integer.class.isAssignableFrom(type) || int.class.isAssignableFrom(type)) {
			return cellVal.intValue();
		} else if (Long.class.isAssignableFrom(type) || long.class.isAssignableFrom(type)) {
			return cellVal.longValue();
		} else if (Double.class.isAssignableFrom(type) || double.class.isAssignableFrom(type)) {
			return cellVal;
		} else if (Float.class.isAssignableFrom(type) || float.class.isAssignableFrom(type)) {
			return cellVal.floatValue();
		} else if (Short.class.isAssignableFrom(type) || short.class.isAssignableFrom(type)) {
			return cellVal.shortValue();
		} else if (Byte.class.isAssignableFrom(type) || byte.class.isAssignableFrom(type)) {
			return cellVal.byteValue();
		} else if (String.class.isAssignableFrom(type)) {
			return cellVal.toString();
		} else if (Boolean.class.isAssignableFrom(type) || boolean.class.isAssignableFrom(type)) {
			return cellVal.intValue() == 1;
		}
		
		return null;
	}

	/**
	 * Read Integer, Long, Double, Short, Byte, Float, Boolean and String from number cell. 
	 * @param cell
	 * @param type
	 * @return
	 */
	protected Object readFromNumberCell(Cell cell, Class<?> type) {
		Double cellVal = cell.getNumericCellValue();
		
		// CellType of cell which contains date is CELL_TYPE_NUMERIC
		if (Time.class.isAssignableFrom(type)) {
			
			Date date = cell.getDateCellValue();
			String time = timeformat.format(date);
			logger.log(Level.INFO, "Time cell value is " + cellVal);
			logger.log(Level.INFO, "Time cell format is " + time);
			return Time.valueOf(time);
		} else if (Date.class.isAssignableFrom(type)) {
			logger.log(Level.INFO, "Date cell value is " + cellVal);
			return cell.getDateCellValue();
		} else if (String.class.isAssignableFrom(type)) {
			return formatNumber(cellVal);
		}
		
		return resolveDouble(cellVal, type);
	}
	
	/**
	 * Read Integer, Long, Double, Short, Byte, Float, Boolean and String from string cell. 
	 * @param cell
	 * @param type
	 * @return
	 */
	@SuppressWarnings("unchecked")
	protected Object readFromStringCell(Cell cell, Class<?> type) {
		String text = cell.getStringCellValue();
		
		if (String.class.isAssignableFrom(type)) {
			return text;
		} else if (Boolean.class.isAssignableFrom(type)) {
			return "yes".equalsIgnoreCase(text) || "true".equalsIgnoreCase(text);
		} else if (Date.class.isAssignableFrom(type)) {
			return DateUtil.resolveDate(text);
		} else if (Integer.class.isAssignableFrom(type)
				|| Long.class.isAssignableFrom(type)
				|| Double.class.isAssignableFrom(type)
				|| Float.class.isAssignableFrom(type)
				|| Short.class.isAssignableFrom(type) 
				|| Byte.class.isAssignableFrom(type)) {
			Double cellVal = new Double(text);
			
			return resolveDouble(cellVal, type);
		} else if (type.isEnum()) {
			if (ICustomEnum.class.isAssignableFrom(type)) {
				return resolveExcelEnum(text, (Class<ICustomEnum>)type);
			} else {
				return resolveEnumValue(text, type);
			}
		}
		
		return null;
	}
	
	/**
	 * Resolve ExcelEnum values.
	 * @param value
	 * @param type
	 * @return
	 */
	protected ICustomEnum resolveExcelEnum(String value, Class<ICustomEnum> type) {
		for (ICustomEnum constant : type.getEnumConstants()) {
			if (constant.value().equalsIgnoreCase(value)) {
				return constant;
			}
		}
		
		return null;
	}
	
	/**
	 * Resolve enum values.
	 * @param value
	 * @param type
	 * @return
	 */
	protected <T> T resolveEnumValue(String value, Class<T> type) {
		for (T constant : type.getEnumConstants()) {
			if (constant.toString().equalsIgnoreCase(value)) {
				return constant;
			}
		}
		
		return null;
	}
	
	public void setValidator(Validator validator) {
		this.validator = validator;
	}

	

}
