package com.linus.excel.validation;

/**
 * 
 * @author lyan2
 */
public class DecimalRangeColumnConstraint extends ColumnConstraint {
	private Double min = Double.MIN_VALUE;
	private Double max = Double.MAX_VALUE;

	public DecimalRangeColumnConstraint() {
		super();
		this.message = "excel.valiation.decimalrange.message";
	}

	@Override
	public boolean isValid(Object value) {
		if (value != null) {
			try {
				Double v = Double.parseDouble(value.toString());
				return (v >= min && v <=max);
			} catch (NumberFormatException e) {
				return false;
			}			
		}
		
		return true;
	}

	public Double getMin() {
		return min;
	}

	public void setMin(Double min) {
		this.min = min;
	}

	public Double getMax() {
		return max;
	}

	public void setMax(Double max) {
		this.max = max;
	}
}
