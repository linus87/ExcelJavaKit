package com.linus.excel;

import java.sql.Time;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;

import com.linus.date.DateUtil;
import com.linus.enums.ICustomEnum;
import com.linus.excel.validation.ExcelValidator;

/**
 * SheetReader.readSheet() support localization messages, default message bundle is "ValidationMessages.properties".
 * 
 * @author lyan2
 */
public class ListSheetReader extends AbstractSheetReader<List<Object>> {

	private final Logger logger = Logger.getLogger(ListSheetReader.class.getName());
	
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

	public List<List<Object>> readSheet(Sheet sheet, int firstRowNum, short firstCellNum, short lastCellNum) {
		
		return readSheet(sheet, firstRowNum, sheet.getLastRowNum(), firstCellNum, lastCellNum);
	}

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
	public List<List<Object>> readSheet(Sheet sheet,
			List<ColumnConfiguration> configs, int firstRowNum, Set<InvalidRowError<List<Object>>> violations) {
		
		return readSheet(sheet, configs, firstRowNum, sheet.getLastRowNum(), violations);
	}
	
	@Override
	public List<List<Object>> readSheet(Sheet sheet,
			List<ColumnConfiguration> configs, int firstRowNum, int lastRowNum,
			Set<InvalidRowError<List<Object>>> violations) {
		if (sheet == null) return null;
		
		ArrayList<List<Object>> list = new ArrayList<List<Object>>();
		ExcelValidator validator = new ExcelValidator();
		
		for (int i = firstRowNum; i < lastRowNum; i++) {
			logger.log(Level.INFO,	"Beginning to read row " + i);
			
			List<Object> obj = readRow(sheet.getRow(i));
			if (obj != null) {
				Set<InvalidCellValueError> errors  = validator.validate(i, obj, configs);
				if (errors == null || errors.size() <= 0) {
					list.add(obj);
				} else {
					InvalidRowError<List<Object>> rowError = new InvalidRowError<List<Object>>(i, obj, null);
					rowError.setCellErrors(errors);
					violations.add(rowError);
					break;
				}
			}
		}
		
		return list;
	}
	
	@Override
	public List<Object> readRow(List<ColumnConfiguration> headers, Row row) {
		
		return readRow(row);
	}

	public List<Object> readRow(Row row) {
		if (row == null) return null;
		
		return readRow(row, (short)0, row.getLastCellNum());
	}

	public List<Object> readRow(Row row, short firstCellNum) {
		if (row == null) return null;
		
		return readRow(row, firstCellNum, row.getLastCellNum());
	}

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
	protected <T> T resolveEnumValue(String value, Class<T> type) {
		for (T constant : type.getEnumConstants()) {
			if (constant.toString().equalsIgnoreCase(value)) {
				return constant;
			}
		}
		
		return null;
	}
	
}
