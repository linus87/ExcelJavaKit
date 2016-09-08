package com.linus.primitive;

public class VariableMain {
	
	// below eight fields have default value 
	static int a;
	static short b;
	static long c;
	static float d;
	static double e;
	static boolean f;
	static char g;
	static byte h;

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// below eight variables don't have have default value
/*		int a;
		short b;
		long c;
		float d;
		double e;
		boolean f;
		char g;
		byte h;*/
		System.out.println(a);
		System.out.println(b);
		System.out.println(c);
		System.out.println(d);
		System.out.println(e);
		System.out.println(f);
		System.out.println(g);
		System.out.println(h);
		
		System.out.println(Long.parseLong(null));
	}

}
