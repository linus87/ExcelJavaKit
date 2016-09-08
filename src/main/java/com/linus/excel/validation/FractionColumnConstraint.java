package com.linus.excel.validation;

public class FractionColumnConstraint extends ColumnConstraint {
	
	private int precision;
	
	public FractionColumnConstraint() {
		super();
		this.message = "excel.valiation.fraction.message";
	}

	/**
	 * Return false only if value is NaN.
	 */
	public boolean isValid(Object value) {
		if (value != null && Double.isNaN((Double)value)) {
			return false;
		}
		
		return true;
	}

	public int getPrecision() {
		return precision;
	}

	public void setPrecision(int precision) {
		this.precision = precision;
	}
	
}
