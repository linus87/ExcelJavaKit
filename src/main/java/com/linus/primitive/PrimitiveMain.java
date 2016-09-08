package com.linus.primitive;

import javax.lang.model.type.PrimitiveType;

public class PrimitiveMain {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		isAssignableFrom();
		convert();
	}
	
	public static void isAssignableFrom() {
		System.out.println("Wrapped classes:");
		System.out.println("Number is Primitive: " + Number.class.isPrimitive());
		System.out.println("Boolean is Primitive: " + Boolean.class.isPrimitive());
		System.out.println("Character is Primitive: " + Character.class.isPrimitive());
		System.out.println("Byte is Primitive: " + Byte.class.isPrimitive());
		System.out.println("Float is Primitive: " + Float.class.isPrimitive());
		System.out.println("Double is Primitive: " + Double.class.isPrimitive());
		System.out.println("Short is Primitive: " + Short.class.isPrimitive());
		System.out.println("Integer is Primitive: " + Integer.class.isPrimitive());
		System.out.println("Long is Primitive: " + Long.class.isPrimitive());
		
		System.out.println("---------------------------------------------------------------------------------------");
		System.out.println("Unwrapped classes:");
		System.out.println("Boolean is Primitive: " + boolean.class.isPrimitive());
		System.out.println("Character is Primitive: " + char.class.isPrimitive());
		System.out.println("Byte is Primitive: " + byte.class.isPrimitive());
		System.out.println("Float is Primitive: " + float.class.isPrimitive());
		System.out.println("Double is Primitive: " + double.class.isPrimitive());
		System.out.println("Short is Primitive: " + short.class.isPrimitive());
		System.out.println("Integer is Primitive: " + int.class.isPrimitive());
		System.out.println("Long is Primitive: " + long.class.isPrimitive());
		
		System.out.println("---------------------------------------------------------------------------------------");
		System.out.println("Unwrapped classes:");
		System.out.println("Boolean is assignable from boolean: " + Boolean.class.isAssignableFrom(boolean.class));
		System.out.println("Character is assignable from char: " + Character.class.isAssignableFrom(char.class));
		System.out.println("Byte is is assignable from byte: " + Byte.class.isAssignableFrom(byte.class));
		System.out.println("Float is is assignable from float: " + Float.class.isAssignableFrom(float.class));
		System.out.println("Double is is assignable from double: " + Double.class.isAssignableFrom(double.class));
		System.out.println("Short is is assignable from short: " + Short.class.isAssignableFrom(short.class));
		System.out.println("Integer is is assignable from int: " + Integer.class.isAssignableFrom(int.class));
		System.out.println("Long is is assignable from long: " + Long.class.isAssignableFrom(long.class));
		
		System.out.println("---------------------------------------------------------------------------------------");
		System.out.println("boolean is assignable from Boolean: " + boolean.class.isAssignableFrom(Boolean.class));
		System.out.println("char is assignable from Character: " + char.class.isAssignableFrom(Character.class));
		System.out.println("byte is is assignable from Byte: " + byte.class.isAssignableFrom(Byte.class));
		System.out.println("float is is assignable from Float: " + float.class.isAssignableFrom(Float.class));
		
		System.out.println("---------------------------------------------------------------------------------------");
		System.out.println("Boolean is assignable from boolean: " + boolean.class.isAssignableFrom(boolean.class));
		System.out.println("Character is assignable from char: " + char.class.isAssignableFrom(char.class));
		System.out.println("Byte is is assignable from byte: " + byte.class.isAssignableFrom(byte.class));
		System.out.println("Float is is assignable from float: " + float.class.isAssignableFrom(float.class));
		
		System.out.println("---------------------------------------------------------------------------------------");
		System.out.println("Primitive Type:");
		System.out.println("PrimitiveType is assignable from boolean: " + PrimitiveType.class.isAssignableFrom(boolean.class));
		System.out.println("Character is assignable from char: " + PrimitiveType.class.isAssignableFrom(char.class));
		System.out.println("PrimitiveType is is assignable from byte: " + PrimitiveType.class.isAssignableFrom(byte.class));
		System.out.println("Float is is assignable from float: " + PrimitiveType.class.isAssignableFrom(float.class));
		
		System.out.println("---------------------------------------------------------------------------------------");
		System.out.println("Primitive Type:");
		System.out.println("PrimitiveType is assignable from boolean: " + PrimitiveType.class.isAssignableFrom(Boolean.class));
		System.out.println("Character is assignable from char: " + PrimitiveType.class.isAssignableFrom(Character.class));
		System.out.println("PrimitiveType is is assignable from byte: " + PrimitiveType.class.isAssignableFrom(Byte.class));
		System.out.println("Float is is assignable from float: " + PrimitiveType.class.isAssignableFrom(Float.class));
		
		System.out.println("---------------------------------------------------------------------------------------");
		System.out.println("Unwrapped classes:");
		System.out.println("Byte: " + Byte.parseByte("58"));
		System.out.println("Byte: " + Byte.parseByte("83"));
	}
	
	public static void convert() {
		double d = 48.25;
		String ds = Double.toString(d);
		ds = ds.substring(0, ds.indexOf("."));
		System.out.println(Integer.parseInt(ds));
		
		System.out.println("a".charAt(0));
		
		char ch = 58;
		System.out.println(ch);
		System.out.println(Character.digit(ch,10));
		System.out.println(Character.forDigit(58,255));
		
		Double doubleO = new Double(38.25);
		System.out.println(doubleO.intValue());
		System.out.println(doubleO.byteValue());
		System.out.println(doubleO.longValue());
		System.out.println(doubleO.floatValue());
		
		doubleO = new Double("hello");
//		System.out.println(doubleO.intValue());
//		System.out.println(doubleO.byteValue());
//		System.out.println(doubleO.longValue());
//		System.out.println(doubleO.floatValue());
	}
	
	
}
