package com.linus.excel;

import java.util.List;
import java.util.Map;

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
public interface ISheetWriter {
	
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
	 * Fill row with a map's values. Display order is the same as it is in list.
	 * @param book
	 * @param sheet
	 * @param row
	 * @param list
	 */
	public void writeRow(Workbook book, Sheet sheet, Row row, List<Object> list);
	
	/**
	 * Fill row with a map's values. Display order is determined by config.getKey() and config.getWriteOrder().
	 * @param book
	 * @param sheet
	 * @param row
	 * @param configs
	 * @param map
	 */
	public void writeRow(Workbook book, Sheet sheet, Row row, List<ColumnConfiguration> configs, Map<String, Object> map);
	
	/**
	 * Create column header title in the first row. Column title is determined by config.getTitle().
	 * @param book
	 * @param sheet
	 * @param configs
	 */
	public void createTitle(Workbook book, Sheet sheet, List<ColumnConfiguration> configs);
	
	/**
	 * Fill sheet with data from list. Argument configs contains the configuration information of each column.
	 * @param book
	 * @param sheet
	 * @param configs
	 * @param list
	 * @param hasTitle
	 */
	public void writeSheet(Workbook book, Sheet sheet, List<ColumnConfiguration> configs, List<Map<String, Object>> list, boolean hasTitle);
	
	/**
	 * Fill sheet with data from list. Argument configs contains the configuration information of each column.
	 * @param book
	 * @param sheet
	 * @param configs
	 * @param list
	 * @param hasTitle
	 */
	public void writeSheet2(Workbook book, Sheet sheet, List<ColumnConfiguration> configs, List<List<Object>> list, boolean hasTitle);
	
	/**
	 * Create freeze pane.
	 * 
	 * @param sheet
	 * @param freezeCols
	 * @param freezeRows
	 */
	public void freeze(Sheet sheet, int freezeCols, int freezeRows);
	
	/**
	 * Hide a column
	 * @param sheet
	 * @param hiddenCol
	 */
	public void hideColumn(Sheet sheet, int hiddenCol);
	
	/**
	 * Set protection password
	 * @param password
	 */
	public void setProtectionPassword(Sheet sheet, String password);
	
}
