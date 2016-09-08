package com.linus.primitive;

/**
 * 
 * @author lyan2
 */
public class StringUtil {

	/**
	 * Repeat a segment of text and return the result.
	 * @param source
	 * @param times
	 * @return
	 */
	public static String repeat(String source, int times) {
		StringBuffer buffer = new StringBuffer();
		while(times > 0) {
			buffer.append(source);
			times--;
		}
		
		return buffer.toString();
	}
}
