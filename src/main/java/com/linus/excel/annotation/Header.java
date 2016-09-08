package com.linus.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.METHOD, ElementType.FIELD})
public @interface Header {
	String title() default "header.default";
	
	/**
	 * Which column does this property get value from excel.
	 * @return
	 */
	int readOrder() default 0;
	
	/**
	 * If this property may be written into excel in the future, set it as true.
	 * @return
	 */
	boolean writable() default true;	
	int writeOrder() default 0;
	
	/**
	 * if user can't modify this cell in excel, set it as false.
	 * @return
	 */
	boolean modifiable() default true;
}