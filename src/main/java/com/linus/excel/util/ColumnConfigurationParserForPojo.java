package com.linus.excel.util;

import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Locale;
import java.util.ResourceBundle;

import com.linus.excel.ColumnConfiguration;
import com.linus.excel.annotation.Header;

public class ColumnConfigurationParserForPojo {
	
	/**
	 * Get column configurations for POJO properties. These configurations are stored in <code>Header</code> annotation.
	 * 
	 * @see Header
	 * 
	 * @param clazz Class
	 * @return A list of column configuration on properties.
	 * @throws IntrospectionException
	 */
	public static ArrayList<ColumnConfiguration> getColumnConfigurations(Class<?> clazz, Locale locale, ResourceBundle bundle) throws IntrospectionException {

		BeanInfo info = Introspector.getBeanInfo(clazz);
		PropertyDescriptor[] descriptors = info.getPropertyDescriptors();
		ArrayList<ColumnConfiguration> configs = new ArrayList<ColumnConfiguration>();

		for (int i = 0; i < descriptors.length; i++) {
			PropertyDescriptor descriptor = descriptors[i];
			Method getter = descriptor.getReadMethod();
			// "Boolean" type property doesn't support "is" getter. Only "boolean" supports. 
			if (getter != null) {
				Header h = descriptor.getReadMethod().getAnnotation(Header.class);
				if (h != null) {
					ColumnConfiguration config = new ColumnConfiguration();
					if (bundle != null && bundle.containsKey(h.title())) {
						config.setTitle(bundle.getString(h.title()));
					} else {
						config.setTitle(h.title());
					}

					config.setKey(descriptor.getName());
					config.setColumnIndex(h.columnIndex());
					config.setWritable(h.writable());
					config.setRawType(h.rawType());
					config.setDisplay(h.display());
					config.setPropertyDescriptor(descriptor);
					configs.add(config);
				}
			}
		}

		return configs;
	}

}
