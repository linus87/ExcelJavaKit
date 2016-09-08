package com.linus.excel;

import javax.validation.ConstraintViolation;
import javax.validation.Path;
import javax.validation.metadata.ConstraintDescriptor;

public class InvalidCellValueError implements ConstraintViolation<Object> {

	public InvalidCellValueError(int rowIndex, int colIndex, Object value, String message) {
		this.rowIndex = rowIndex;
		this.colIndex = colIndex;
		this.cellValue = value;
		this.message = message;
	}	
	
	public int getColIndex() {
		return colIndex;
	}

	public void setColIndex(int colIndex) {
		this.colIndex = colIndex;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	public Object getCellValue() {
		return cellValue;
	}

	public void setCellValue(Object cellValue) {
		this.cellValue = cellValue;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}

	private int colIndex;
	private int rowIndex;
	private Object cellValue;
	private String message;
	public String getMessageTemplate() {
		// TODO Auto-generated method stub
		return null;
	}

	public Object getRootBean() {
		// TODO Auto-generated method stub
		return null;
	}

	public Class<Object> getRootBeanClass() {
		// TODO Auto-generated method stub
		return null;
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
		return null;
	}

	public Object getInvalidValue() {
		// TODO Auto-generated method stub
		return null;
	}

	public ConstraintDescriptor<?> getConstraintDescriptor() {
		// TODO Auto-generated method stub
		return null;
	}

	public <U> U unwrap(Class<U> type) {
		// TODO Auto-generated method stub
		return null;
	}

	
}
