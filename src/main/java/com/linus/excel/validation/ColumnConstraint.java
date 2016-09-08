package com.linus.excel.validation;


/**
 * 
 * @author lyan2
 */
public abstract class ColumnConstraint {
	protected String message = "excel.validation.message";
	protected Boolean requried;
	protected Class<?> type;
	
	/**
	 * Validate this constraint.
	 * @return
	 */
	public abstract boolean isValid(Object value);
	
	public boolean equal(Object obj1, Object obj2) {
		if (obj1 != null && obj2 != null) {
			return obj1.equals(obj2);
		}
		
		return false;
	}

	/**
	 * After get the message template from bundle, it's still need to replace the arguments in the message template.
	 * @param message
	 * @return
	 */
	public String resolveMessage(String message) {
		return message;
	}
	
	protected Class<?> getType() {
		return type;
	}

	public Boolean getRequried() {
		return requried;
	}

	public void setRequried(Boolean requried) {
		this.requried = requried;
	}

	public void setType(Class<?> type) {
		this.type = type;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}
}
