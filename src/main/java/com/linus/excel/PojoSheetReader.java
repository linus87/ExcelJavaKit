package com.linus.excel;

import java.beans.IntrospectionException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.validation.ConstraintViolation;
import javax.validation.Validator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;

import com.linus.excel.util.ColumnConfigurationParserForJson;

public class PojoSheetReader<T> extends AbstractSheetReader<T> {
	
	private final Logger logger = Logger.getLogger(PojoSheetReader.class.getName());
	public static Map<Type, ArrayList<ColumnConfiguration>> sheetHeaders = new ConcurrentHashMap<Type, ArrayList<ColumnConfiguration>>();
	
	private Validator validator;
	protected Class<T> clazz;
	
	public List<T> readSheet(Sheet sheet, Class<T> clazz, int firstDataRow, Set<InvalidRowError<T>> constraintViolations) {
		ArrayList<ColumnConfiguration> headers = null;

		if (sheetHeaders.containsKey(clazz)) {
			 headers = sheetHeaders.get(clazz);
		} else {
			try {
				headers = ColumnConfigurationParserForJson.getColumnConfigurations(clazz);
			} catch (IntrospectionException e) {
				logger.log(Level.SEVERE, "Failed to get column configuration from class - " + clazz.getName() + ", due to bean instropection exception.");
			}
			
			sheetHeaders.put(clazz, headers);
		}
		
		this.clazz = clazz;
		return readSheet(sheet, headers, firstDataRow, constraintViolations);
	}
	
	@Override
	public List<T> readSheet(Sheet sheet, List<ColumnConfiguration> headers, int firstDataRow, Set<InvalidRowError<T>> constraintViolations) {
		return readSheet(sheet, headers, firstDataRow, sheet.getLastRowNum(), constraintViolations);
	}
	
	@Override
	public List<T> readSheet(Sheet sheet, List<ColumnConfiguration> headers, int firstDataRow, int lastRowNum, Set<InvalidRowError<T>> constraintViolations) {
		if (headers == null || sheet == null) return null;
		
		ArrayList<T> list = new ArrayList<T>();
		
		for (int i = firstDataRow; i < lastRowNum; i++) {
			logger.log(Level.INFO,	"Beginning to read row " + i);
			Row row = sheet.getRow(i);
			
			if (row != null) {
				T obj = readRow(headers, row);
				if (obj != null) {
					if (validator != null) {
						Set<ConstraintViolation<T>> violations = validator.validate(obj);
						
						if (violations == null || violations.isEmpty()) {
							list.add(obj);
						} else {
							if (constraintViolations != null) {
								InvalidRowError<T> rowError = new InvalidRowError<T>(i, obj, "Failed to read from row " + i);
								rowError.setCellErrors(transferConstraintViolation(i, violations));
								constraintViolations.add(rowError);
								break;
							}
						}
					} else {
						list.add(obj);
					}
				}
			}
		}
		
		return list;

	}
	
	private Set<InvalidCellError> transferConstraintViolation(int rowNum, Set<ConstraintViolation<T>> violations) {
		Set<InvalidCellError> errors = new HashSet<InvalidCellError>(violations.size());
		Iterator<ConstraintViolation<T>> iterator = violations.iterator();
		while (iterator.hasNext()) {
			ConstraintViolation<T> violation = iterator.next();
			InvalidCellError error = new InvalidCellError(rowNum, violation.getPropertyPath().toString(), violation.getInvalidValue(), violation.getMessage());
			errors.add(error);
		}
		
		return errors;
	}
	
	/**
	 * Read a single row and convert it into specified type instance. All cell values will be instance properties. 
	 */
	@Override
	public T readRow(List<ColumnConfiguration> headers, Row row) {
		
		T o = null;

		if (headers != null && !headers.isEmpty()) {
			try {
				o = clazz.newInstance();
			} catch (InstantiationException e) {
				logger.log(Level.WARNING, "Class " + clazz.getName()
						+ " can't be instantiated!");
				e.printStackTrace();
				return o;
			} catch (IllegalAccessException e) {
				logger.log(Level.WARNING, "Class " + clazz.getName()
						+ " can't be instantiated! Because it's not accssable.");
				e.printStackTrace();
				return o;
			}

			Iterator<ColumnConfiguration> iter = headers.iterator();
			ColumnConfiguration header = null;
			while (iter.hasNext()) {
				header = iter.next();
				
				Cell cell = row.getCell(header.getColumnIndex(), MissingCellPolicy.RETURN_NULL_AND_BLANK);
				Method setter = header.getPropertyDescriptor().getWriteMethod();
				Object value = readCell(cell, header.getPropertyDescriptor().getReadMethod().getReturnType());

				logger.log(Level.INFO, "Property "
						+ header.getPropertyDescriptor().getName()
						+ " get value " + value);
				try {
					if (value != null ) {
						setter.invoke(o, value);
					}
				} catch (IllegalAccessException | IllegalArgumentException
						| InvocationTargetException e) {
					logger.log(Level.WARNING, "Property "
							+ header.getPropertyDescriptor().getName()
							+ " can't be set on Class " + clazz.getName()
							+ " instance, value is " + value);
					e.printStackTrace();
				}
			}
		}
		return o;
	}
	
	public void setValidator(Validator validator) {
		this.validator = validator;
	}

}
