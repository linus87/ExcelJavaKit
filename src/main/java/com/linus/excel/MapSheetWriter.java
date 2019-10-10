package com.linus.excel;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.linus.excel.validation.ColumnConstraint;
import com.linus.excel.validation.IntegerRangeColumnConstraint;
import com.linus.excel.validation.RangeColumnConstraint;

/**
 * 
 * @author lyan2
 */
public class MapSheetWriter extends AbstractSheetWriter<Map<String, Object>> {
	private final Logger logger = Logger.getLogger(MapSheetWriter.class.getName());
	private Map<Integer, CellStyle> styleMapping = new HashMap<Integer, CellStyle>();

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
	}

	public void createSubHead(Workbook book, Sheet sheet, List<ColumnConfiguration> configs) {
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
	
	/**
	 * Create a row filled with sample data.
	 * 
	 * @param book
	 * @param sheet
	 * @param configs
	 */
	public void createSampleRow(Workbook book, Sheet sheet, List<ColumnConfiguration> configs) {
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

}
