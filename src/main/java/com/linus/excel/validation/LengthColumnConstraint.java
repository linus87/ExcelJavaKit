package com.linus.excel.validation;

/**
 * Validate if text is too long.
 * 
 * @author lyan2
 */
public class LengthColumnConstraint extends ColumnConstraint {

	private int length;
	
	public LengthColumnConstraint(int maxlength) {
		super();
		this.length = maxlength;
		this.message = "excel.validation.length.message";
	}

	public boolean isValid(Object value) {
		String text = null;
		
		if (null != value && !(value instanceof String)) {
			text = value.toString();
		} else {
			text = (String) value;
		}

		if (text != null && text.length() > length) {
			return false;
		}

		return true;
	}
	
	public String resolveMessage(String message) {
		if (message != null) {
			message = message.replaceAll("\\{length\\}", String.valueOf(this.getLength()));
		}
		return message;
	}

	public void setLength(int length) {
		this.length = length;
	}

	public int getLength() {
		return length;
	}

}
