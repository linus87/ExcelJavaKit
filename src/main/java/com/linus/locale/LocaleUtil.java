package com.linus.locale;

import java.util.Locale;

public class LocaleUtil {
	/**
	 * Get locale string format like: en_US, zh_CN.
	 * @param locale
	 * @return
	 */
	public static String getLocale(Locale locale) {
//		return locale.getLanguage() + "_" + locale.getCountry();
		return locale.getCountry();
	}
}
