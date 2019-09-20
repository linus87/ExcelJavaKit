package com.linus.excel;

import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.validation.ConstraintViolation;
import javax.validation.Validator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 
 * @author lyan2
 */
public interface ISheetReader {
	
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
	 * Get cells values from the the first cell to the last cell which is physical defined. They are stored in a list by the order they appears in a row.
	 * Note: This method only get cells values from the first column to the last cell which is physical defined(cell has value or cell type).
	 * 
	 * @see ISheetReader#readRow(Row, short, short)
	 * 
	 * @param row
	 * @return List of cell values
	 */
	public List<Object> readRow(Row row);
	
	/**
	 * Get cells values from the specified first column to the last cell which is physical defined(cell has value or cell type). 
	 * These values are stored in a list by the order they appears in a row.
	 * 
	 * @see ISheetReader#readRow(Row, short, short)
	 * 
	 * @param row
	 * @param firstCellNum The number of the first cell to read.
	 * @return List of cell values
	 */
	public List<Object> readRow(Row row, short firstCellNum);
	
	/**
	 * Read cell values and store them in a list by the order they appears in a row.
	 * Note: If cell is blank or it doesn't exist, null is returned.
	 * @param row
	 * @param firstCellNum The number of the first cell to read.
	 * @param lastCellNum The number of the last cell to read.
	 * @return List of cell values
	 */
	public List<Object> readRow(Row row, short firstCellNum, short lastCellNum);
	
	/**
	 * Each column maps a attribute, so we get values from cells and store them as attributes.
	 * The first argument <code>headers</code> store the columns configurations, for example, which column maps to which attribute.
	 * 
	 * @param headers
	 * @param row
	 * @return
	 */
	public Map<String, Object> readRow(List<ColumnConfiguration> headers, Row row);
	
	/**
	 * Read all rows from the first row, cells of each row should be stored in a List by the order they appears in a row.
	 * This method should read the whole sheet from the specified row to the last row.
	 * @param sheet
	 * @param clazz
	 * @param firstRowNum The number of the first row to read.
	 * @param violations
	 * @return
	 */
	public List<List<Object>> readSheet(Sheet sheet, int firstRowNum);
	
	/**
	 * Read all rows from the first row, cells of each row should be stored in a List by the order they appears in a row.
	 * This method should read the whole sheet from the specified row to the last row. But only read specified columns.
	 * @param sheet
	 * @param firstRowNum The number of the first row to read.
	 * @param firstCellNum The number of the first cell to read.
	 * @param lastCellNum The number of the last cell to read. lastCellNum must bigger than or equal to firstCellNum.
	 * @param violations
	 * @return
	 */
	public List<List<Object>> readSheet(Sheet sheet, int firstRowNum, short firstCellNum, short lastCellNum);
	
	/**
	 * Read a few rows from the specified first row to the specified last row. For each row, only read from specified first column to the specified last column.
	 * Cells of each row should be stored in a List by the order they appears in a row.
	 * This method should read the whole sheet from the specified row to the specified row. But only read specified columns.
	 * @param sheet
	 * @param firstRowNum The number of the first row to read.
	 * @param lastRowNum The number of the last row to read. lastRowNum must bigger than or equal to firstRowNum.
	 * @param firstCellNum The number of the first cell to read. lastCellNum must bigger than or equal to firstCellNum.
	 * @param lastCellNum The number of the last cell to read.
	 * @param violations
	 * @return
	 */
	public List<List<Object>> readSheet(Sheet sheet, int firstRowNum, int lastRowNum, short firstCellNum, short lastCellNum);
	
	/**
	 * Read all rows from the first row, cells of each row should be stored in a List by the order they appears in a row.
	 * Each row is also validated according to ColumnConfigurations. Reading will stop when a invalid row is read. 
	 * @param sheet
	 * @param clazz
	 * @param firstRowNum The number of the first row to read.
	 * @param violations Validation errors will be stored here.
	 * @return
	 */
	public List<Map<String, Object>> readSheet(Sheet sheet, List<ColumnConfiguration> headers, int firstRowNum, Set<ConstraintViolation<Object>> violations);
	
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
	public List<Map<String, Object>> readSheet(Sheet sheet, List<ColumnConfiguration> headers, int firstRowNum, int lastRowNum, Set<ConstraintViolation<Object>> violations);
	
	/**
	 * Read sheet, and each row will be represented as a array.
	 * Each row is also validated according to ColumnConfigurations. Reading will stop when a invalid row is read. 
	 * @param sheet
	 * @param headers
	 * @param firstRowNum The number of the first row to read.
	 * @param violations Validation errors will be stored here.
	 * @return
	 */
	public List<List<Object>> readSheet2(Sheet sheet, List<ColumnConfiguration> headers, int firstRowNum, Set<ConstraintViolation<Object>> violations);
	
	/**
	 * Read sheet, and each row will be represented as a array.
	 * Each row is also validated according to ColumnConfigurations. Reading will stop when a invalid row is read. 
	 * @param sheet
	 * @param headers
	 * @param firstRowNum
	 * @param lastRowNum
	 * @param violations
	 * @return
	 */
	public List<List<Object>> readSheet2(Sheet sheet, List<ColumnConfiguration> headers, int firstRowNum, int lastRowNum, Set<ConstraintViolation<Object>> violations);

	/**
	 * Read cells from a row and store cells value in a specified class's instance as its properties.
	 * 
	 * @param headers
	 * @param row
	 * @param clazz
	 * @return
	 */
	public <T> T readRow(List<ColumnConfiguration> headers, Row row, Class<T> clazz);
	
	/**
	 * If a row is not valid, read operation will stop and argument violations will be filled with ConstraintViolations.
	 * 
	 * If argument violations is null, this method should read all rows and return all valid objects. But if argument violations is not null,
	 * reading operation should stops when it encounters an invalid object, and violations set will be filled with ConstraintViolations objects.
	 * However, those former valid objects should be returned.
	 * 
	 * @param sheet
	 * @param clazz
	 * @param firstDataRow
	 * @param violations
	 * @return
	 */
	public <T> List<T> readSheet(Sheet sheet, Class<T> clazz, int firstDataRow, Set<ConstraintViolation<T>> violations);
		
	public void setValidator(Validator validator);
}
