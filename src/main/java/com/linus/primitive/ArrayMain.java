package com.linus.primitive;

import java.lang.reflect.Array;

public class ArrayMain {
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		System.out.println();
		
		System.out.println("---------------------------------------------------------------------------------------");
		System.out.println("Unwrapped classes:");
		
		int[] array =  (int[]) Array.newInstance(int.class, 8);
		System.out.println(array.length);
		
		Object array2 =  Array.newInstance(int.class, 8);
		System.out.println(Array.getLength(array2));
		
		System.out.println("Byte: " + Byte.parseByte("83"));
	}
	
}
