package com.linus.excel.validation;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.ResourceBundle;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.linus.excel.ColumnConfiguration;
import com.linus.excel.InvalidRowError;

/**
 * Validate data from a row of excel sheet. For each sheet reading, you need to create a new ExcelValidator object.
 * 
 * @author lyan2
 */
public class ListValidator {
	private Logger logger = Logger.getLogger(ListValidator.class.getName());
	private String bundleBaseName = "ExcelValidationMessages";
	private Locale locale;
	private ResourceBundle bundle;
	
	/**
	 * Validate excel row data.
	 * @param rowIndex
	 * @param map
	 * @param configs
	 * @return
	 */
	public Set<InvalidRowError<List<Object>>> validate(int rowIndex, List<Object> list, List<ColumnConfiguration> configs) {
		Set<InvalidRowError<List<Object>>> errors = new HashSet<InvalidRowError<List<Object>>>();
		
		if (configs != null && !configs.isEmpty() && list != null) {
			for (ColumnConfiguration config : configs) {
				Object value = list.get(config.getColumnIndex());
				List<ColumnConstraint> constraints = config.getConstraints();
				for (ColumnConstraint constraint : constraints) {
					try {
						if (!constraint.isValid(value)) {
							logger.warning("validation: " + constraint.getMessage());
							String message = getBundle().getString(constraint.getMessage());
							String invalidMessage = getBundle().getString("excel.validation.invalidcell.message");
							invalidMessage = invalidMessage.replaceFirst("\\{row\\}", String.valueOf(rowIndex + 1));
							invalidMessage = invalidMessage.replaceFirst("\\{title\\}", config.getTitle());
							invalidMessage += constraint.resolveMessage(message);
							errors.add(new InvalidRowError<List<Object>>(rowIndex, config.getColumnIndex(), value, invalidMessage));
						}
					} catch(ClassCastException e) {
						logger.log(Level.WARNING, e.getMessage());
					}
					
				}
			}
		}
		
		return errors;
	}
	
	/**
	 * Validate excel row data.
	 * @param rowIndex
	 * @param map
	 * @param configs
	 * @return
	 */
	public List<InvalidRowError<List<Object>>> validate2(int rowIndex, List<Object> list, List<ColumnConfiguration> configs) {
		List<InvalidRowError<List<Object>>> errors = new ArrayList<InvalidRowError<List<Object>>>();
		int hiddenColumnNums = 1;
		
		if (configs != null && !configs.isEmpty() && list != null) {
			for (ColumnConfiguration config : configs) {
				Object value = list.get(config.getColumnIndex());
				List<ColumnConstraint> constraints = config.getConstraints();
				for (ColumnConstraint constraint : constraints) {
					try {
						if (!constraint.isValid(value)) {
							logger.warning("validation: " + constraint.getMessage());
							String message = getBundle().getString(constraint.getMessage());
							String invalidMessage = getBundle().getString("excel.validation.invalidcell.message");
							invalidMessage = invalidMessage.replaceFirst("\\{row\\}", String.valueOf(rowIndex + 1));
							invalidMessage = invalidMessage.replaceFirst("\\{column\\}", String.valueOf(config.getColumnIndex() + hiddenColumnNums));
							invalidMessage += constraint.resolveMessage(message);
							errors.add(new InvalidRowError<List<Object>>(rowIndex, config.getColumnIndex(), value, invalidMessage));
						}
					} catch(ClassCastException e) {
						logger.log(Level.WARNING, e.getMessage());
					}
					
				}
				
				if (!config.getDisplay()) {
					hiddenColumnNums++;
				}
			}
		}
		
		return errors;
	}
	
	public Locale getLocale() {
		return locale == null ? Locale.getDefault() : locale;
		
		// in spring, use this code
//		return locale == null ? LocaleContextHolder.getLocale() : locale;
	}
	
	public void setLocale(Locale locale) {
		this.locale = locale;
	}

	/**
	 * Default return "ExcelValidationMessages" resource bundle.
	 * @return
	 */
	public ResourceBundle getBundle() {
		return bundle == null ? bundle = ResourceBundle.getBundle(bundleBaseName) : bundle;
	}

	public void setBundle(ResourceBundle bundle) {
		this.bundle = bundle;
	}

	/**
	 *  Default return "ExcelValidationMessages" resource bundle.
	 * @return
	 */
	public String getBundleBaseName() {
		return bundleBaseName;
	}

	public void setBundleBaseName(String bundleBaseName) {
		this.bundleBaseName = bundleBaseName;
	}
	
}
