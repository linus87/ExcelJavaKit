package com.linus.excel.validation;

/**
 * 
 * @author lyan2
 */
public class IntegerRangeColumnConstraint extends ColumnConstraint {
	private Integer min = Integer.MIN_VALUE;
	private Integer max = Integer.MAX_VALUE;

	public IntegerRangeColumnConstraint() {
		super();
		this.message = "excel.valiation.integerrange.message";
	}

	@Override
	public boolean isValid(Object value) {
		if (value != null) {
			try {
				Integer v = Integer.parseInt(value.toString());
				return (v >= min && v <=max);
			} catch (NumberFormatException e) {
				return false;
			}			
		}
		
		return true;
	}

	public Integer getMin() {
		return min;
	}

	public void setMin(Integer min) {
		this.min = min;
	}

	public Integer getMax() {
		return max;
	}

	public void setMax(Integer max) {
		this.max = max;
	}
}
