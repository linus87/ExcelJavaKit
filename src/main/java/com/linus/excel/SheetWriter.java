package com.linus.excel;

import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.linus.date.DateUtil;
import com.linus.excel.validation.ColumnConstraint;
import com.linus.excel.validation.DoubleColumnConstraint;
import com.linus.excel.validation.IntegerRangeColumnConstraint;
import com.linus.excel.validation.RangeColumnConstraint;
import com.linus.primitive.StringUtil;

/**
 * 
 * @author lyan2
 */
public class SheetWriter implements ISheetWriter {
	private final Logger logger = Logger.getLogger(SheetWriter.class.getName());
	private Map<Integer, CellStyle> styleMapping = new HashMap<Integer, CellStyle>();
	private int firstDataRowNum = 0;

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

	@Override
	public void writeRow(Workbook book, Sheet sheet, Row row, List<Object> list) {
		CellStyle cellStyle = book.createCellStyle();
		cellStyle.setFont(defaultFont);
		
		int column = 0;
		for (Object value : list) {			
			createCell(book, sheet, row, column++, value, cellStyle);
		}
	}

	@Override
	public void writeRow(Workbook book, Sheet sheet, Row row, List<ColumnConfiguration> configs, Map<String, Object> map) {
		for (ColumnConfiguration config : configs) {
			if (config != null) {
				CellStyle cellStyle = styleMapping.get(config.getColumnIndex());
				if (cellStyle == null) {
					cellStyle = book.createCellStyle();
					cellStyle.setLocked(!config.getWritable());
					cellStyle.setFont(defaultFont);
					cellStyle.setWrapText(true);
					styleMapping.put(config.getColumnIndex(), cellStyle);
				}
				createCell(book, sheet, row, config, map.get(config.getKey()), cellStyle);
			}
		}
	}

	@Override
	public void writeSheet(Workbook book, Sheet sheet, List<ColumnConfiguration> configs,
			List<Map<String, Object>> list, boolean hasTitle) {
		if (hasTitle)
			createTitle(book, sheet, configs);
		createSampleRow(book, sheet, configs);

		int rowNum = firstDataRowNum;
		for (Map<String, Object> map : list) {
			Row row = sheet.createRow(rowNum++);
			writeRow(book, sheet, row, configs, map);
		}

		for (ColumnConfiguration config : configs) {
			if (!config.getDisplay()) {
				hideColumn(sheet, config.getColumnIndex());
			}
		}

		if (firstDataRowNum > sheet.getLastRowNum()) {
			// no data
			return;
		}

		for (ColumnConfiguration config : configs) {
			List<ColumnConstraint> constraints = config.getConstraints();
			for (ColumnConstraint constraint : constraints) {
				if (constraint instanceof RangeColumnConstraint) {
					// only support single
					if (((RangeColumnConstraint) constraint).isAllowMultiple()) break;
					
					XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
					XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
					    dvHelper.createExplicitListConstraint(((RangeColumnConstraint) constraint).getPickList());

					CellRangeAddressList addressList = new CellRangeAddressList(firstDataRowNum,  sheet.getLastRowNum(), config.getColumnIndex(), config.getColumnIndex());
					XSSFDataValidation validation =(XSSFDataValidation)dvHelper.createValidation(dvConstraint, addressList);

					// Display pick list when user click the cell.
					validation.setSuppressDropDownArrow(true);

					// Note this extra method call. If this method call is
					// omitted, or if the
					// boolean value false is passed, then Excel will not
					// validate the value the
					// user enters into the cell.
					validation.setShowErrorBox(((RangeColumnConstraint) constraint).getMustInRange());
					sheet.addValidationData(validation);
					break;
				}

				if (constraint instanceof IntegerRangeColumnConstraint) {
					IntegerRangeColumnConstraint iconstraint = (IntegerRangeColumnConstraint) constraint;
					XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
					XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
							.createNumericConstraint(XSSFDataValidationConstraint.ValidationType.INTEGER,
									XSSFDataValidationConstraint.OperatorType.BETWEEN, "=" + iconstraint.getMin(), "="
											+ iconstraint.getMax());

					CellRangeAddressList addressList = new CellRangeAddressList(firstDataRowNum, sheet.getLastRowNum(),
							config.getColumnIndex(), config.getColumnIndex());
					XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint,
							addressList);

					// Display pick list when user click the cell.
					validation.setSuppressDropDownArrow(true);

					// Note this extra method call. If this method call is
					// omitted, or if the
					// boolean value false is passed, then Excel will not
					// validate the value the
					// user enters into the cell.
					validation.setShowErrorBox(true);
					sheet.addValidationData(validation);
					break;
				}
			}
		}

		adjustColumnWidth(sheet, configs);
	}

	@Override
	public void writeSheet2(Workbook book, Sheet sheet, List<ColumnConfiguration> configs, List<List<Object>> list,
			boolean hasTitle) {
		if (hasTitle)
			createTitle(book, sheet, configs);

		int rowNum = firstDataRowNum;
		for (List<Object> obj : list) {
			Row row = sheet.createRow(rowNum);
			writeRow(book, sheet, row, obj);
			rowNum++;
		}

	}

	@Override
	public void createTitle(Workbook book, Sheet sheet, List<ColumnConfiguration> configs) {
		Row row = sheet.createRow(firstDataRowNum++);
		CellStyle headerStyle = book.createCellStyle();
		headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
		headerStyle.setFillForegroundColor(IndexedColors.LIME.getIndex());
		headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		headerStyle.setFont(titleFont);
		headerStyle.setWrapText(true);

		for (ColumnConfiguration config : configs) {
			if (config != null) {
				createCell(book, sheet, row, config.getColumnIndex(), config.getTitle(), headerStyle);
			}
		}

		createSubHead(book, sheet, configs);
	}

	private void createSubHead(Workbook book, Sheet sheet, List<ColumnConfiguration> configs) {
		Row row = sheet.createRow(firstDataRowNum++);
		CellStyle headerStyle = book.createCellStyle();
		headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
		headerStyle.setFillForegroundColor(IndexedColors.LIME.getIndex());
		headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		headerStyle.setFont(this.defaultFont);
		headerStyle.setWrapText(true);

		for (ColumnConfiguration config : configs) {
			if (config != null) {
				createCell(book, sheet, row, config.getColumnIndex(), config.getLabel(), headerStyle);
			}
		}
	}

	/**
	 * Freeze some columns and rows.
	 * 
	 * @param sheet
	 * @param freezeRows
	 * @param freezeCols
	 * @param password
	 */
	public void freeze(Sheet sheet, int freezeCols, int freezeRows) {
		sheet.createFreezePane(freezeCols, freezeRows);
	}

	/**
	 * Hide a column.
	 * 
	 * @param sheet
	 * @param hiddenCol
	 */
	public void hideColumn(Sheet sheet, int hiddenCol) {
		sheet.setColumnHidden(hiddenCol, true);
	}

	public int getFirstDataRowNum() {
		return firstDataRowNum;
	}

	public void setFirstDataRowNum(int firstDataRowNum) {
		this.firstDataRowNum = firstDataRowNum;
	}

	@Override
	public void setProtectionPassword(Sheet sheet, String password) {
		sheet.protectSheet(password);
	}

	/**
	 * Create general cell and in String type.
	 * 
	 * @param book
	 * @param row
	 * @param config
	 * @param value
	 */
	private void createCell(Workbook book, Row row, ColumnConfiguration config, Object value, CellStyle style) {
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

		if (value != null) {
			cell.setCellValue(value.toString());
		} else {
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

	/**
	 * Create a row filled with sample data.
	 * 
	 * @param book
	 * @param sheet
	 * @param configs
	 */
	private void createSampleRow(Workbook book, Sheet sheet, List<ColumnConfiguration> configs) {
		Row row = sheet.createRow(firstDataRowNum++);

		// sample data is locked;
		CellStyle style = book.createCellStyle();
		style.setLocked(true);
		style.setWrapText(true);
		style.setFont(defaultFont);

		for (ColumnConfiguration config : configs) {
			if (config != null) {
				createCell(book, row, config, config.getSample(), style);
			}
		}
	}

	/**
	 * <p>
	 * Adjust column width. If there is visible character length setting, then
	 * use this length setting. If there is no length setting, adjusts the
	 * column width to fit the contents.
	 * 
	 * To compute the actual number of visible characters, Excel uses the
	 * following formula (Section 3.3.1.12 of the OOXML spec):
	 * </p>
	 * <code>
	 *     width = Truncate([{Number of Visible Characters} *
	 *      {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
	 * </code>
	 * <p>
	 * The maximum column width for an individual cell is 255 characters.
	 * </p>
	 * 
	 * @param sheet
	 * @param configs
	 */
	protected void adjustColumnWidth(Sheet sheet, List<ColumnConfiguration> configs) {

		for (ColumnConfiguration config : configs) {

			if (config.getLength() != null) {
				// character length + 2, 256 is a character's width;
				int columnWidth = 512 * (config.getLength() + 2);
				sheet.setColumnWidth(config.getColumnIndex(), columnWidth);
			} else {
				sheet.autoSizeColumn(config.getColumnIndex());
			}

		}
	}

	private Font defaultFont;
	private Font titleFont;

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
