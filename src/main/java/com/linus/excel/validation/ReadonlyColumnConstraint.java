package com.linus.excel.validation;

/**
 * 
 * @author lyan2
 */
public class ReadonlyColumnConstraint extends ColumnConstraint {
	protected Object primitiveValue;

	public ReadonlyColumnConstraint() {
		super();
		this.message = "excel.validation.readonly.message";
	}

	public boolean isValid(Object value) {
		if (this.primitiveValue != null) {
			return equal(primitiveValue, value);
		}
		
		return true;
	}
	
	public Object getPrimitiveValue() {
		return primitiveValue;
	}

	public void setPrimitiveValue(Object primitiveValue) {
		this.primitiveValue = primitiveValue;
	}
	
}
