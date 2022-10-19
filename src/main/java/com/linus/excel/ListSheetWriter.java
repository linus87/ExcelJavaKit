package com.linus.excel;

import java.util.List;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
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
    
	public ListSheetWriter(Workbook book, List<ColumnConfiguration> configs) {
        super(book, configs);
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
	public void writeSheet(Workbook book, Sheet sheet, List<List<Object>> list,
			boolean hasTitle) {
		if (hasTitle)
			createTitle(book, sheet, configs);

		int rowNum = firstDataRowNum;
		for (List<Object> array : list) {
			Row row = sheet.createRow(rowNum);
			writeRow(book, sheet, row, array);
			rowNum++;
		}

	}

	@Override
	public void createTitle(Workbook book, Sheet sheet, List<ColumnConfiguration> configs) {
		Row row = sheet.createRow(firstDataRowNum++);
		CellStyle headerStyle = this.getHeaderCellStyle();

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
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setFillForegroundColor(IndexedColors.LIME.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		headerStyle.setFont(this.defaultFont);
		headerStyle.setWrapText(true);

		for (ColumnConfiguration config : configs) {
			if (config != null) {
				createCell(book, sheet, row, config.getColumnIndex(), config.getLabel(), headerStyle);
			}
		}
	}

	private Font defaultFont;

	public Font getDefaultFont() {
		return defaultFont;
	}

	public void setDefaultFont(Font defaultFont) {
		this.defaultFont = defaultFont;
	}

}
