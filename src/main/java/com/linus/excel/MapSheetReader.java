package com.linus.excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;

import com.linus.excel.validation.ExcelValidator;

/**
 * SheetReader.readSheet() support localization messages, default message bundle is "ValidationMessages.properties".
 * 
 * @author lyan2
 */
public class MapSheetReader extends AbstractSheetReader<Map<String, Object>> {

	private final Logger logger = Logger.getLogger(MapSheetReader.class.getName());
	
	@Override
	public List<Map<String, Object>> readSheet(Sheet sheet,
			List<ColumnConfiguration> headers, int firstRowNum,
			Set<InvalidRowError<Map<String, Object>>> violations) {
		if (sheet == null) return null;
		
		return readSheet(sheet, headers, firstRowNum, sheet.getLastRowNum(), violations);
	}

	@Override
	public List<Map<String, Object>> readSheet(Sheet sheet,
			List<ColumnConfiguration> configs, int firstRowNum, int lastRowNum,
			Set<InvalidRowError<Map<String, Object>>> violations) {
		if (sheet == null) return null;
		
		ArrayList<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		ExcelValidator validator = new ExcelValidator();
		
		for (int i = firstRowNum; i <= lastRowNum; i++) {
			logger.log(Level.INFO,	"Beginning to read row " + i);
			
			Map<String, Object> obj = readRow(configs, sheet.getRow(i));
			if (obj != null) {
				Set<InvalidCellError> errors  = validator.validate(i, obj, configs);
				if (errors == null || errors.size() <= 0) {
					list.add(obj);
				} else {
					InvalidRowError<Map<String, Object>> rowError = new InvalidRowError<Map<String, Object>>(i, obj, "Failed to read from row " + i);
					rowError.setCellErrors(errors);
					violations.add(rowError);
					break;
				}
			}
		}
		
		return list;
	}
	
	@Override
	public Map<String, Object> readRow(List<ColumnConfiguration> headers, Row row) {
		if (row == null) return null;
		
		Map<String, Object> map = new HashMap<String, Object>();
		Iterator<ColumnConfiguration> iter = headers.iterator();
		
		while (iter.hasNext()) {
			ColumnConfiguration header = iter.next();
			
			// if cell doesn't exist, return null;
			Cell cell = row.getCell(header.getColumnIndex(), MissingCellPolicy.CREATE_NULL_AS_BLANK);
			Object value = null;
			
			if (header.getType() != null) {
				value = readCell(cell, header.getType());
			} else {
				if (!"attachment".equalsIgnoreCase(header.getRawType())) {
					value = readCell(cell);
				}
			}		
			
			map.put(header.getKey(), value);
		}
		
		return map;
	}

}
