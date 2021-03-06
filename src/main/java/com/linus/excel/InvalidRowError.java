package com.linus.excel;

import javax.validation.ConstraintViolation;
import javax.validation.Path;
import javax.validation.metadata.ConstraintDescriptor;

public class InvalidRowError<T> implements ConstraintViolation<T> {
	
	private int rowIndex;
	private int colIndex;
	private Object cellValue;
	private Object value;
	private Path propertyPath;
	private String message;

	public InvalidRowError(int rowIndex, int colIndex, Object cellValue, String message) {
		this.rowIndex = rowIndex;
		this.colIndex = colIndex;
		this.cellValue = cellValue;
		this.message = message;
	}
	
	public InvalidRowError(int rowIndex, Object bean, Path property, String message) {
		this.rowIndex = rowIndex;
		this.value = bean;
		this.propertyPath = property;
		this.message = message;
	}
	
	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}

	public String getMessageTemplate() {
		// TODO Auto-generated method stub
		return null;
	}

	public T getRootBean() {
		return (T)value;
	}

	public Class<T> getRootBeanClass() {
		// TODO Auto-generated method stub
		return (Class<T>) value.getClass();
	}

	public Object getLeafBean() {
		// TODO Auto-generated method stub
		return null;
	}

	public Object[] getExecutableParameters() {
		// TODO Auto-generated method stub
		return null;
	}

	public Object getExecutableReturnValue() {
		// TODO Auto-generated method stub
		return null;
	}

	public Path getPropertyPath() {
		// TODO Auto-generated method stub
		return propertyPath;
	}

	public Object getInvalidValue() {
		return value != null ? value : cellValue;
	}

	public ConstraintDescriptor<?> getConstraintDescriptor() {
		// TODO Auto-generated method stub
		return null;
	}

	public <U> U unwrap(Class<U> type) {
		// TODO Auto-generated method stub
		return null;
	}

	public int getColIndex() {
		return colIndex;
	}

	public void setColIndex(int colIndex) {
		this.colIndex = colIndex;
	}

	public Object getCellValue() {
		return cellValue;
	}

	public void setCellValue(Object cellValue) {
		this.cellValue = cellValue;
	}
}
