package com.linus.excel;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.sql.Time;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;

import com.linus.date.DateUtil;
import com.linus.enums.ICustomEnum;

public abstract class AbstractSheetReader<T> implements ISheetReader<T> {
	private final Logger logger = Logger.getLogger(AbstractSheetReader.class.getName());
	
	protected SimpleDateFormat timeformat = new SimpleDateFormat("HH:mm:ss");
	
	@SuppressWarnings("unchecked")
	public static <T extends Number> T parseNumber(String text, Class<T> targetClass) {
		String trimmed = text.trim();

		if (Byte.class == targetClass) {
			return (T) (isHexNumber(trimmed) ? Byte.decode(trimmed) : Byte.valueOf(trimmed));
		}
		else if (Short.class == targetClass) {
			return (T) (isHexNumber(trimmed) ? Short.decode(trimmed) : Short.valueOf(trimmed));
		}
		else if (Integer.class == targetClass) {
			return (T) (isHexNumber(trimmed) ? Integer.decode(trimmed) : Integer.valueOf(trimmed));
		}
		else if (Long.class == targetClass) {
			return (T) (isHexNumber(trimmed) ? Long.decode(trimmed) : Long.valueOf(trimmed));
		}
		else if (BigInteger.class == targetClass) {
			return (T) (isHexNumber(trimmed) ? decodeBigInteger(trimmed) : new BigInteger(trimmed));
		}
		else if (Float.class == targetClass) {
			return (T) Float.valueOf(trimmed);
		}
		else if (Double.class == targetClass) {
			return (T) Double.valueOf(trimmed);
		}
		else if (BigDecimal.class == targetClass || Number.class == targetClass) {
			return (T) new BigDecimal(trimmed);
		}
		else {
			throw new IllegalArgumentException(
					"Cannot convert String [" + text + "] to target class [" + targetClass.getName() + "]");
		}
	}
	
	public Object readCell(Cell cell) {
		if (cell == null) return null;
		
		switch (cell.getCellType()) {
		case BLANK: return null;
		case ERROR: return null;
		case NUMERIC:
			return cell.getNumericCellValue();
		case STRING:
			return cell.getStringCellValue();
		case BOOLEAN: 
			return cell.getBooleanCellValue();
		case FORMULA:
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
		case BLANK: return null;
		case ERROR: return null;
		case NUMERIC:
			return readFromNumberCell(cell, type);
		case STRING:
			return readFromStringCell(cell, type);
		case BOOLEAN: 
			Boolean value = cell.getBooleanCellValue();
			if (Boolean.class.isAssignableFrom(type)) {
				return value;
			} else {
				return value.toString();
			}
		case FORMULA:
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
	protected Object resolveNumber(String cellVal, Class<?> type) {
		Double  value = new Double(cellVal);
		
		if (Integer.class.isAssignableFrom(type) || int.class.isAssignableFrom(type)) {
			return value.intValue();
		} else if (Long.class.isAssignableFrom(type) || long.class.isAssignableFrom(type)) {
			return value.longValue();
		} else if (Double.class.isAssignableFrom(type) || double.class.isAssignableFrom(type)) {
			return value;
		} else if (Float.class.isAssignableFrom(type) || float.class.isAssignableFrom(type)) {
			return value.floatValue();
		} else if (Short.class.isAssignableFrom(type) || short.class.isAssignableFrom(type)) {
			return value.shortValue();
		} else if (Byte.class.isAssignableFrom(type) || byte.class.isAssignableFrom(type)) {
			return value.byteValue();
		} else if (BigInteger.class.isAssignableFrom(type)) {
			return new BigInteger(cellVal);
		} else if (BigDecimal.class.isAssignableFrom(type)) {
			return new BigDecimal(cellVal);
		} else if (String.class.isAssignableFrom(type)) {
			return value.toString();
		} else if (Boolean.class.isAssignableFrom(type) || boolean.class.isAssignableFrom(type)) {
			return value.intValue() == 1;
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
		} else if (Calendar.class.isAssignableFrom(type)) {
			logger.log(Level.INFO, "Date cell value is " + cellVal);
			return new Calendar.Builder().setInstant(cell.getDateCellValue()).build();
		} else if (String.class.isAssignableFrom(type)) {
			return formatNumber(cellVal);
		}
		
		return resolveNumber(cellVal.toString(), type);
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
		} else if (Calendar.class.isAssignableFrom(type)) {
			return new Calendar.Builder().setInstant(DateUtil.parseDate(text)).build();
		} else if (Integer.class.isAssignableFrom(type)
				|| Long.class.isAssignableFrom(type)
				|| Double.class.isAssignableFrom(type)
				|| Float.class.isAssignableFrom(type)
				|| Short.class.isAssignableFrom(type) 
				|| Byte.class.isAssignableFrom(type)) {
			
			return resolveNumber(text, type);
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
	
	/**
	 * Determine whether the given {@code value} String indicates a hex number,
	 * i.e. needs to be passed into {@code Integer.decode} instead of
	 * {@code Integer.valueOf}, etc.
	 */
	private static boolean isHexNumber(String value) {
		int index = (value.startsWith("-") ? 1 : 0);
		return (value.startsWith("0x", index) || value.startsWith("0X", index) || value.startsWith("#", index));
	}
	
	/**
	 * Decode a {@link java.math.BigInteger} from the supplied {@link String} value.
	 * <p>Supports decimal, hex, and octal notation.
	 * @see BigInteger#BigInteger(String, int)
	 */
	private static BigInteger decodeBigInteger(String value) {
		int radix = 10;
		int index = 0;
		boolean negative = false;

		// Handle minus sign, if present.
		if (value.startsWith("-")) {
			negative = true;
			index++;
		}

		// Handle radix specifier, if present.
		if (value.startsWith("0x", index) || value.startsWith("0X", index)) {
			index += 2;
			radix = 16;
		}
		else if (value.startsWith("#", index)) {
			index++;
			radix = 16;
		}
		else if (value.startsWith("0", index) && value.length() > 1 + index) {
			index++;
			radix = 8;
		}

		BigInteger result = new BigInteger(value.substring(index), radix);
		return (negative ? result.negate() : result);
	}
}
