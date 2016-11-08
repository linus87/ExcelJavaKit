package com.linus.excel.validation;

public class FractionColumnConstraint extends ColumnConstraint {
	
	private int precision;
	
	public FractionColumnConstraint() {
		super();
		this.message = "excel.validation.fraction.message";
	}

	/**
	 * Return false only if value is NaN.
	 */
	public boolean isValid(Object value) {
		
		if (value != null) {
			Double dValue = null;
			try {
				dValue = Double.valueOf(value.toString());
			} catch (Exception e) {
				return false;
			}
			
			if (dValue != null) {
				String str = value.toString();
				int dotIndex = str.indexOf(".") + 1;
				if (dotIndex > 0 && dotIndex + precision < str.length()) {
					while (dotIndex < str.length()) {
						if (str.charAt(dotIndex) != '0') {
							return false;
						}
						dotIndex++;
					}
				}
			}
		}
		
		return true;
	}

	public int getPrecision() {
		return precision;
	}

	public void setPrecision(int precision) {
		this.precision = precision;
	}
	
	public static void main(String[] args) {
		FractionColumnConstraint constraint = new FractionColumnConstraint();
		constraint.setPrecision(0);
		System.out.println(constraint.isValid("34.883"));
		System.out.println(constraint.isValid("34.8"));
		System.out.println(constraint.isValid("34"));
		System.out.println(constraint.isValid("34."));
		System.out.println(constraint.isValid("33453"));
		System.out.println(constraint.isValid(""));
		System.out.println(constraint.isValid(34.8));
		System.out.println(constraint.isValid(34.883));
		System.out.println(constraint.isValid(34d));
		System.out.println(constraint.isValid(34.));
	}
	
}
