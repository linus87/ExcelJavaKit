package com.linus.excel.validation;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 
 * @author lyan2
 */
public class PatternColumnConstraint extends ColumnConstraint {
	private Pattern pattern;

	public PatternColumnConstraint(Pattern pattern) {
		super();
		this.pattern = pattern;
		this.message = "excel.valiation.pattern.message";
	}

	public boolean isValid(Object value) {
		if (value != null) {
			String text = (String) value;
			Matcher matcher = pattern.matcher(text);
			
			if (text != null && !matcher.matches()) {
				return false;
			}
		}
		
		return true;
	}

	public Pattern getPattern() {
		return pattern;
	}

	public void setPattern(Pattern pattern) {
		this.pattern = pattern;
	}
	
}
