package com.linus.excel;

import java.util.List;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 
 * @author lyan2
 */
public class ListSheetWriter extends AbstractSheetWriter<List<Object>> {
	private final Logger logger = Logger.getLogger(ListSheetWriter.class.getName());
	private int firstDataRowNum = 0;

	@Override
	public void writeRow(Workbook book, Sheet sheet, Row row, List<ColumnConfiguration> configs, List<Object> list) {
		CellStyle cellStyle = book.createCellStyle();
		cellStyle.setFont(defaultFont);
		
		int column = 0;
		for (Object value : list) {			
			createCell(book, sheet, row, column++, value, cellStyle);
		}
	}

	@Override
	public void writeSheet(Workbook book, Sheet sheet, List<ColumnConfiguration> configs, List<List<Object>> list,
			boolean hasTitle) {
		if (hasTitle)
			createTitle(book, sheet, configs);

		int rowNum = firstDataRowNum;
		for (List<Object> array : list) {
			Row row = sheet.createRow(rowNum);
			writeRow(book, sheet, row, configs, array);
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
