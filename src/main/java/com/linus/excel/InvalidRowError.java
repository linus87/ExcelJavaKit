package com.linus.excel;

import java.util.Set;

import javax.validation.ConstraintViolation;
import javax.validation.Path;
import javax.validation.metadata.ConstraintDescriptor;

public class InvalidRowError<T> implements ConstraintViolation<T> {
	
	private int rowIndex;
	private T value;
	private String message;
	private Set<InvalidCellError> cellErrors;

	public InvalidRowError(int rowIndex, T value, String message) {
		this.rowIndex = rowIndex;
		this.value = value;
		this.message = message;
	}
	
	public Set<InvalidCellError> getCellErrors() {
		return cellErrors;
	}

	public void setCellErrors(Set<InvalidCellError> cellErrors) {
		this.cellErrors = cellErrors;
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
		return value;
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
		return null;
	}

	public Object getInvalidValue() {
		// TODO Auto-generated method stub
		return value;
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
