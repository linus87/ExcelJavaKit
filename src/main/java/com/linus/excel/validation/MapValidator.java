package com.linus.excel.validation;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.Set;

import com.linus.excel.ColumnConfiguration;
import com.linus.excel.InvalidRowError;

public class MapValidator {
	private String bundleBaseName = "ExcelValidationMessages";
	private ResourceBundle bundle;
	
	/**
	 * Validate excel row data. Error message will tell us which row and which column (with column title) has error.
	 * @param rowIndex
	 * @param map
	 * @param configs
	 * @return
	 */
	public Set<InvalidRowError<Map<String, Object>>> validate(int rowIndex, Map<String, Object> map, List<ColumnConfiguration> configs) {
		Set<InvalidRowError<Map<String, Object>>> errors = new HashSet<InvalidRowError<Map<String, Object>>>();
		
		if (configs != null && !configs.isEmpty() && map != null) {
			for (ColumnConfiguration config : configs) {
				Object value = map.get(config.getKey());
				List<ColumnConstraint> constraints = config.getConstraints();
				for (ColumnConstraint constraint : constraints) {
					if (!constraint.isValid(value)) {
						String message = getBundle().getString(constraint.getMessage());
						String invalidMessage = getBundle().getString("excel.validation.invalidcell.message");
						invalidMessage = invalidMessage.replaceFirst("\\{row\\}", String.valueOf(rowIndex + 1));
						invalidMessage = invalidMessage.replaceFirst("\\{title\\}", config.getTitle());
						invalidMessage += constraint.resolveMessage(message);
						errors.add(new InvalidRowError<Map<String, Object>>(rowIndex, config.getColumnIndex(), value, invalidMessage));
					}
				}
			}
		}
		
		return errors;
	}
	
	/**
	 * Validate excel row data. Error message will tell us which row and which column (in number) has error.
	 * @param rowIndex
	 * @param map
	 * @param configs
	 * @return
	 */
	public List<InvalidRowError<Map<String, Object>>> validate2(int rowIndex, Map<String, Object> map, List<ColumnConfiguration> configs) {
		List<InvalidRowError<Map<String, Object>>> errors = new ArrayList<InvalidRowError<Map<String, Object>>>();
		int hiddenColumnNums = 0;
		
		if (configs != null && !configs.isEmpty() && map != null) {
			for (ColumnConfiguration config : configs) {
				Object value = map.get(config.getKey());
				List<ColumnConstraint> constraints = config.getConstraints();
				for (ColumnConstraint constraint : constraints) {
					if (!constraint.isValid(value)) {
						String message = getBundle().getString(constraint.getMessage());
						String invalidMessage = getBundle().getString("excel.validation.invalidcell.message");
						invalidMessage = invalidMessage.replaceFirst("\\{row\\}", String.valueOf(rowIndex + 1));
						invalidMessage = invalidMessage.replaceFirst("\\{column\\}", String.valueOf(config.getColumnIndex() + hiddenColumnNums));
						invalidMessage += constraint.resolveMessage(message);
						errors.add(new InvalidRowError<Map<String, Object>>(rowIndex, config.getColumnIndex(), value, invalidMessage));
					}
				}
				
				if (!config.getDisplay()) {
					hiddenColumnNums++;
				}
			}
		}
		
		return errors;
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
