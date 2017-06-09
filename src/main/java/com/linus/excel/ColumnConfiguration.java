package com.linus.excel;

import java.beans.PropertyDescriptor;
import java.util.ArrayList;
import java.util.List;

import com.linus.excel.validation.ColumnConstraint;

/**
 * Store each column's configuration information.
 * @author lyan2
 */
public class ColumnConfiguration {
	
	/**
	 * Excel column header title. In fact, it's local label.
	 */
	private String title;
	
	private String label;
	
	/**
	 * Character length.
	 */
	private Integer length;
	
	/**
	 * Which excel column to read data.
	 */
	private int readOrder;
	
	/**
	 * Which excel row to read data.
	 */
	private int writeOrder;
	
	/**
	 * Whether user can change the default value in the cell.
	 */
	private Boolean writable;
	
	/**
	 * The PropertyDescriptor object of the property which the cell maps to.
	 */
	private PropertyDescriptor propertyDescriptor;
	
	private String rawType;
	
	private Class<?> type;
	
	// for JSON and XML conversion
	private String key;
	
	// sample data
	private String sample;
	
	private List<ColumnConstraint> constraints = new ArrayList<ColumnConstraint>();
	
	/**
	 * hide this column or not
	 */
	private Boolean display = true;
	
	public String getTitle() {
		return title;
	}
	public void setTitle(String title) {
		this.title = title;
	}
	
	public Integer getLength() {
	    return length;
	}
	public void setLength(Integer length) {
	    this.length = length;
	}
	public Boolean getWritable() {
		return writable;
	}
	public void setWritable(Boolean writable) {
		this.writable = writable;
	}
	public PropertyDescriptor getPropertyDescriptor() {
		return propertyDescriptor;
	}
	public void setPropertyDescriptor(PropertyDescriptor propertyDescriptor) {
		this.propertyDescriptor = propertyDescriptor;
	}

	public int getReadOrder() {
		return readOrder;
	}
	public void setReadOrder(int readOrder) {
		this.readOrder = readOrder;
	}
	public int getWriteOrder() {
		return writeOrder;
	}
	public void setWriteOrder(int writeOrder) {
		this.writeOrder = writeOrder;
	}

	public String getRawType() {
		return rawType;
	}
	public void setRawType(String rawType) {
		this.rawType = rawType;
	}

	public String getKey() {
		return key;
	}
	public void setKey(String key) {
		this.key = key;
	}
	public List<ColumnConstraint> getConstraints() {
		return constraints;
	}
	public void setConstraints(List<ColumnConstraint> constraints) {
		this.constraints = constraints;
	}

	public Boolean getDisplay() {
		return display;
	}
	public void setDisplay(Boolean display) {
		this.display = display;
	}

	public Class<?> getType() {
		return type;
	}
	public void setType(Class<?> type) {
		this.type = type;
	}
	public String getSample() {
		return sample;
	}
	public void setSample(String sample) {
		this.sample = sample;
	}
	public String getLabel() {
		return label;
	}
	public void setLabel(String label) {
		this.label = label;
	}

}
