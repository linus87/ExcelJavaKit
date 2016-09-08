package com.linus.excel.validation;

import java.util.HashMap;
import java.util.Map;

/**
 * 
 * @author lyan2
 */
public class UniqueColumnConstraint extends ColumnConstraint {
	private Map<Object, Object> existedValues = new HashMap<Object, Object>();

	public UniqueColumnConstraint() {
		super();
		this.message = "excel.validation.unique.message";
	}

	@Override
	public boolean isValid(Object value) {
		if (existedValues.containsKey(value)) {
			return false;
		}
		
		existedValues.put(value, null);
		return true;
	}
	
	/**
	 * Clear this constraint.
	 */
	public void clearAll() {
		existedValues.clear();
	}

}
