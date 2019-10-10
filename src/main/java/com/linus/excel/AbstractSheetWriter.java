package com.linus.excel;

import java.util.Calendar;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.linus.date.DateUtil;
import com.linus.excel.util.StringUtil;
import com.linus.excel.validation.ColumnConstraint;
import com.linus.excel.validation.DoubleColumnConstraint;

public abstract class AbstractSheetWriter<T> implements ISheetWriter<T> {
	
	private final Logger logger = Logger.getLogger(AbstractSheetWriter.class.getName());
	
	protected int firstDataRowNum = 0;
	
	@Override
	public void createCell(Workbook book, Sheet sheet, Row row, ColumnConfiguration config, Object value,
			CellStyle style) {
		if (config == null)
			return;

		CellStyle cellStyle = book.createCellStyle();
		cellStyle.cloneStyleFrom(style);

		if (config.getRawType() == null) {
			// attachment doesn't have raw type.
			createCell(book, row, config, value, cellStyle);
			return;
		}

		switch (config.getRawType().toUpperCase()) {
		case "DOUBLE":
			createDoubleCell(book, row, config, value, cellStyle);
			break;
		case "DATE":
			createDateCell(book, row, config, value, cellStyle);
			break;
		case "DATETIME":
			createDateTimeCell(book, row, config, value, cellStyle);
			break;
		case "TIME":
			createTimeCell(book, row, config, value, cellStyle);
			break;
		case "PERCENT":
			createPercentCell(book, row, config, value, cellStyle);
			break;
		case "COMBOBOX":
		case "PICKLIST":
			cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		case "STRING":
		case "TEXTAREA":
		default:
			createCell(book, row, config, value, cellStyle);
			break;
		}

	}

	@Override
	public void createCell(Workbook book, Sheet sheet, Row row, int column, Object value, CellStyle style) {
		CellStyle cellStyle = book.createCellStyle();
		cellStyle.cloneStyleFrom(style);
		Cell cell = null;

		if (value instanceof Number) {
			cell = row.createCell(column, Cell.CELL_TYPE_NUMERIC);
			cellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
			cell.setCellValue(((Number) value).doubleValue());
			cell.setCellStyle(cellStyle);
		} else if (value instanceof Date) {
			cell = row.createCell(column, Cell.CELL_TYPE_NUMERIC);
			cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
			cell.setCellStyle(cellStyle);
			cell.setCellValue((Date) value);
		} else if (value instanceof Boolean) {
			cell = row.createCell(column, Cell.CELL_TYPE_BOOLEAN);
			cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
			cell.setCellStyle(cellStyle);
			cell.setCellValue((Boolean) value);
		} else if (value == null) {
			cell = row.createCell(column, Cell.CELL_TYPE_BLANK);
		} else {
			cell = row.createCell(column, Cell.CELL_TYPE_STRING);
			cell.setCellValue(value.toString());
			cell.setCellStyle(cellStyle);
		}

	}
	
	/**
	 * Create general cell and in String type.
	 * 
	 * @param book
	 * @param row
	 * @param config
	 * @param value
	 */
	protected void createCell(Workbook book, Row row, ColumnConfiguration config, Object value, CellStyle style) {
		Cell cell = row.createCell(config.getColumnIndex(), Cell.CELL_TYPE_STRING);
		if (value != null) {
			cell.setCellValue(value.toString());
		} else {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		}

		style.setDataFormat(HSSFDataFormat.getBuiltinFormat("text"));
		cell.setCellStyle(style);
	}

	/**
	 * Create cell in percentage format, percentage display in format:0.00%.
	 * 
	 * @param book
	 * @param row
	 * @param config
	 * @param value
	 */
	private void createPercentCell(Workbook book, Row row, ColumnConfiguration config, Object value, CellStyle style) {
		Cell cell = row.createCell(config.getColumnIndex(), Cell.CELL_TYPE_NUMERIC);
		
		style.setAlignment(CellStyle.ALIGN_CENTER);
		DataFormat df = book.createDataFormat();
		style.setDataFormat(df.getFormat("0.00%"));
		cell.setCellStyle(style);
		
		if (value != null && value instanceof Number) {
			cell.setCellValue(((Number)value).doubleValue());
		} else if (value != null) {
			cell.setCellValue(value.toString());
		}  else {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		}
		
	}

	/**
	 * Horizontal align datetime right. Decimal digits are determined by
	 * DoubleColumnConstraint.getDigits().
	 * 
	 * if value is LinkedHashMap object, it's treated like {value:#.##,
	 * currency:'RMB'}.
	 * 
	 * @param book
	 * @param row
	 * @param config
	 * @param value
	 */
	private void createDoubleCell(Workbook book, Row row, ColumnConfiguration config, Object value, CellStyle style) {
		Cell cell = row.createCell(config.getColumnIndex(), Cell.CELL_TYPE_NUMERIC);
		style.setAlignment(CellStyle.ALIGN_RIGHT);

		String currency = null;
		String format = null;

		if(value instanceof LinkedHashMap) {
			currency = (String)((LinkedHashMap<?, ?>) value).get("currency");
			value = ((LinkedHashMap<?, ?>) value).get("value");
		}

		for (ColumnConstraint constraint : config.getConstraints()) {
			if (constraint instanceof DoubleColumnConstraint) {
				format = "0";
				if (((DoubleColumnConstraint) constraint).getDigits() > 0) {
					format += "." + StringUtil.repeat("0", ((DoubleColumnConstraint) constraint).getDigits());
				}
				style.setDataFormat(book.createDataFormat().getFormat(format));
			}
		}

		if (currency != null) {
			if (format == null) {
				format = "0.\"(" + currency + ")\"";
			} else {
				format += "\"(" + currency + ")\"";
			}
		}

		if (format != null) {
			style.setDataFormat(book.createDataFormat().getFormat(format));
		}
		cell.setCellStyle(style);

		if (value != null) {
			if (value instanceof Number) {
				cell.setCellValue(((Number) value).doubleValue());
			} else {
				try {
					cell.setCellValue(Double.parseDouble(value.toString()));
				} catch (Exception e) {
					logger.log(Level.WARNING, "Double cell value is not a valid double number.");
				}
			}
		} else {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		}
	}

	/**
	 * Horizontal align datetime center, and display datetime in
	 * format:yyyy-mm-dd.
	 * 
	 * @param book
	 * @param row
	 * @param config
	 * @param value
	 */
	private void createDateCell(Workbook book, Row row, ColumnConfiguration config, Object value, CellStyle style) {
		Cell cell = row.createCell(config.getColumnIndex(), Cell.CELL_TYPE_NUMERIC);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		DataFormat df = book.createDataFormat();
		style.setDataFormat(df.getFormat("yyyy-MM-dd"));
		cell.setCellStyle(style);

		if (value != null) {
			if (value instanceof Date) {
				cell.setCellValue(DateUtil.formatISODate((Date) value, null));
			} else if (value instanceof String) {
				cell.setCellValue((String) value);
			} else if (value instanceof Number) {
				Calendar date = Calendar.getInstance();
				date.setTimeInMillis((long) value);
				cell.setCellValue(DateUtil.formatISODate(date.getTime(), null));
			}
		} else {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		}
	}

	/**
	 * Horizontal align datetime center, and display datetime in
	 * format:yyyy-mm-dd hh:mm:ss.
	 * 
	 * @param book
	 * @param row
	 * @param config
	 * @param value
	 */
	private void createDateTimeCell(Workbook book, Row row, ColumnConfiguration config, Object value, CellStyle style) {
		Cell cell = row.createCell(config.getColumnIndex(), Cell.CELL_TYPE_NUMERIC);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		DataFormat df = book.createDataFormat();
		style.setDataFormat(df.getFormat("yyyy-MM-dd HH:mm:ss"));
		cell.setCellStyle(style);

		if (value instanceof Date && value != null) {
			cell.setCellValue(DateUtil.formatISODateTime((Date) value, null));
		} else if (value instanceof String && value != null) {
			cell.setCellValue((String) value);
		} else if (value instanceof Number) {
			Calendar date = Calendar.getInstance();
			date.setTimeInMillis((long) value);
			cell.setCellValue(DateUtil.formatISODateTime(date.getTime(), null));
		} else {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		}
	}

	/**
	 * Horizontal align time center, and display time in format:hh:mm:ss.
	 * 
	 * @param book
	 * @param row
	 * @param config
	 * @param value
	 */
	private void createTimeCell(Workbook book, Row row, ColumnConfiguration config, Object value, CellStyle style) {
		Cell cell = row.createCell(config.getColumnIndex(), Cell.CELL_TYPE_NUMERIC);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		DataFormat df = book.createDataFormat();
		style.setDataFormat(df.getFormat("HH:mm:ss"));
		cell.setCellStyle(style);

		if (value instanceof Date && value != null) {
			cell.setCellValue(DateUtil.formatTime((Date) value));
		} else if (value instanceof String && value != null) {
			cell.setCellValue((String) value);
		} else {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		}
	}

	protected Font defaultFont;
	protected Font titleFont;

	public Font getDefaultFont() {
		return defaultFont;
	}

	public void setDefaultFont(Font defaultFont) {
		this.defaultFont = defaultFont;
	}

	public Font getTitleFont() {
		return titleFont;
	}

	public void setTitleFont(Font titleFont) {
		this.titleFont = titleFont;
	}
}
