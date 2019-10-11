package com.linus.excel;

import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 
 * @author lyan2
 */
public interface ISheetReader<T> {
	
	/**
	 * Read raw cell value. 
	 * Note: Blank or Error cell return null.
	 * @param cell
	 * @return May return Double, String, Date, Boolean or null.
	 */
	public Object readCell(Cell cell);
	
	/**
	 * Read cell value and convert it into specified type value.
	 * It should support String, Boolean, Date, Time, Double, Float, Long, Integer, Short, Byte, Enum and ICustomEnum.
	 * @param cell
	 * @param type
	 * @return
	 */
	public Object readCell(Cell cell, Class<?> type);
	
	/**
	 * Each column maps a attribute, so we get values from cells and store them as attributes.
	 * The first argument <code>headers</code> store the columns configurations, for example, which column maps to which attribute.
	 * 
	 * @param headers
	 * @param row
	 * @return
	 */
	public T readRow(List<ColumnConfiguration> headers, Row row);
	
	/**
	 * Read all rows from the first row, cells of each row should be stored in a List by the order they appears in a row.
	 * Each row is also validated according to ColumnConfigurations. Reading will stop when a invalid row is read. 
	 * @param sheet
	 * @param clazz
	 * @param firstRowNum The number of the first row to read.
	 * @param violations Validation errors will be stored here.
	 * @return
	 */
	public List<T> readSheet(Sheet sheet, List<ColumnConfiguration> headers, int firstRowNum, Set<InvalidRowError<T>> violations);
	
	/**
	 * Read sheet, and each row will be represented as a JSON object.
	 * Each row is also validated according to ColumnConfigurations. Reading will stop when a invalid row is read. 
	 * @param sheet
	 * @param headers
	 * @param firstRowNum The number of the first row to read.
	 * @param lastRowNum The number of the last row to read.
	 * @param violations Validation errors will be stored here.
	 * @return
	 */
	public List<T> readSheet(Sheet sheet, List<ColumnConfiguration> headers, int firstRowNum, int lastRowNum, Set<InvalidRowError<T>> violations);
	
}
