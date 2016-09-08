package com.linus.excel.validation;

public class NotNullColumnConstraint extends ColumnConstraint {
	
	public NotNullColumnConstraint() {
		super();
		this.message = "excel.validation.notnull.message";
	}
	
	public boolean isValid(Object value) {
		if (value == null) {
			return false;
		}
		
		if (value instanceof String) {
			return ! ((String) value).isEmpty();
		}
		
		return true;
	}
}
