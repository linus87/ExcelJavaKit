package com.linus.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.METHOD, ElementType.FIELD})
public @interface Header {
	String title() default "header.default";
	
	int columnIndex() default 0;
	
	/**
	 * if user can't modify this cell in excel, set it as false.
	 * @return
	 */
	boolean writable() default true;
		
	String rawType() default "STRING";
	
	boolean display() default true;
}