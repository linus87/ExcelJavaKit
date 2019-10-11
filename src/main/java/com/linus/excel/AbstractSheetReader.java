package com.linus.excel;

import java.sql.Time;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.validation.Validator;

import org.apache.poi.ss.usermodel.Cell;

import com.linus.date.DateUtil;
import com.linus.enums.ICustomEnum;

public abstract class AbstractSheetReader<T> implements ISheetReader<T> {
	private final Logger logger = Logger.getLogger(AbstractSheetReader.class.getName());
	
	protected SimpleDateFormat timeformat = new SimpleDateFormat("HH:mm:ss");
	
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
	 * Remove scientific notation
	 * @param value
	 * @return
	 */
	protected String formatNumber(Number value) {
		DecimalFormat df = new DecimalFormat("0");
		return df.format(value);
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
		} else if (Time.class.isAssignableFrom(type)) {
			return DateUtil.parseTime(text);
		} else if (Date.class.isAssignableFrom(type)) {
			return DateUtil.parseDate(text);
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
	protected <P> P resolveEnumValue(String value, Class<P> type) {
		for (P constant : type.getEnumConstants()) {
			if (constant.toString().equalsIgnoreCase(value)) {
				return constant;
			}
		}
		
		return null;
	}
	
}
