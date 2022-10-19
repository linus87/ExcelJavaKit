package com.linus.excel;

import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * This interface contains two kinds of write APIs: One to write each row with a list, another with a map's values. For list, the write order is
 * the same as element's order in list. For map, the write order is determined by ColumnConfiguration.getWriteOrder().
 * 
 * @author lyan2
 */
public interface ISheetWriter<T> {
	
	/**
	 * Create a new cell in row and set cell type according to value type.
	 * @param book
	 * @param sheet
	 * @param row
	 * @param column
	 * @param value
	 * @param style Cell style.
	 */
	public void createCell(Workbook book, Sheet sheet, Row row, int column, Object value, CellStyle style);
	
	/**
	 * Create a new cell in row, column is specified by config.getWriteOrder(). Cell type is determined by config.getRawType().
	 * @param book
	 * @param sheet
	 * @param row
	 * @param config
	 * @param value
	 */
	public void createCell(Workbook book, Sheet sheet, Row row, ColumnConfiguration config, Object value, CellStyle style);

	/**
	 * Fill row with a map's values. Display order is determined by config.getKey() and config.getWriteOrder().
	 * @param book
	 * @param sheet
	 * @param row
	 * @param configs
	 * @param map
	 */
	public void writeRow(Workbook book, Sheet sheet, Row row, T map);
	
	/**
	 * Fill sheet with data from list. Argument configs contains the configuration information of each column.
	 * @param book
	 * @param sheet
	 * @param configs
	 * @param list
	 * @param hasTitle
	 */
	public void writeSheet(Workbook book, Sheet sheet, List<T> list, boolean hasTitle);
	
	/**
	 * Create freeze pane.
	 * 
	 * @param sheet
	 * @param freezeCols
	 * @param freezeRows
	 */
	public default void freeze(Sheet sheet, int freezeCols, int freezeRows) {
		sheet.createFreezePane(freezeCols, freezeRows);
	}
	
	/**
	 * Hide a column
	 * @param sheet
	 * @param hiddenCol
	 */
	public default void hideColumn(Sheet sheet, int hiddenCol) {
		sheet.setColumnHidden(hiddenCol, true);
	}
	
	/**
	 * Set protection password
	 * @param password
	 */
	public default void setProtectionPassword(Sheet sheet, String password) {
		sheet.protectSheet(password);
	}
	
}
