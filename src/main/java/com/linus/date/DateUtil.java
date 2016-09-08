package com.linus.date;

import java.sql.Time;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * All kinds of date operation.
 * 
 * @author lyan2
 */
public class DateUtil {
	private static final Logger logger = Logger.getLogger(DateUtil.class.getName());
	
	// Database time format
	private static final String TIME_FORMAT = "HH:mm:ss";
	
	// Database datetime format
	private static final String ISO_DATETIME_FORMAT = "yyyy-MM-dd HH:mm:ss";
	
	// Excel date format may be following three types.
	private static final String DOT_DATE_FORMAT = "yyyy.MM.dd";
	private static final String SLASH_DATE_FORMAT = "yyyy/MM/dd";
	private static final String ISO_DATE_FORMAT = "yyyy-MM-dd";
	
	// CS API and SaleForce use this time format, in fact, it's ISO 8601 Date and time format.'Z' means UTC time.
	private static final String UTC_DATETIME_FORMAT = "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'";

	// Chinese bei jing time zone
	public final static TimeZone BEIJING_TIMEZONE = TimeZone.getTimeZone("GMT+8:00");

	private static ThreadLocal<DateFormat> timeThread = new ThreadLocal<DateFormat>();
	private static ThreadLocal<DateFormat> simpleDateThread = new ThreadLocal<DateFormat>();
	private static ThreadLocal<DateFormat> slashDateThread = new ThreadLocal<DateFormat>();
	private static ThreadLocal<DateFormat> isoDateThread = new ThreadLocal<DateFormat>();
	private static ThreadLocal<DateFormat> isoDateTimeThread = new ThreadLocal<DateFormat>();
	private static ThreadLocal<DateFormat> isoDateTimeZoneThread = new ThreadLocal<DateFormat>();
	
	/**
	 * Get date format for: HH:mm:ss
	 * @return
	 */
	public static DateFormat getTimeFormat() {
		DateFormat df = timeThread.get();
		if (df == null) {
			df = new SimpleDateFormat(TIME_FORMAT);
			timeThread.set(df);
		}

		return df;
	}

	/**
	 * DateFormat for parsing date like: yyyy.MM.dd
	 * @return
	 */
	public static DateFormat getDotDateFormat() {
		DateFormat df = simpleDateThread.get();
		if (df == null) {
			df = new SimpleDateFormat(DOT_DATE_FORMAT);
			simpleDateThread.set(df);
		}

		return df;
	}
	
	/**
	 * DateFormat for parsing date like: yyyy/MM/dd
	 * @return
	 */
	public static DateFormat getSlashDateFormat() {
		DateFormat df = slashDateThread.get();
		if (df == null) {
			df = new SimpleDateFormat(SLASH_DATE_FORMAT);
			slashDateThread.set(df);
		}

		return df;
	}
	
	/**
	 * DateFormat for parsing date like: yyyy-MM-dd
	 * @return
	 */
	public static DateFormat getISODateFormat() {
		DateFormat df = isoDateThread.get();
		if (df == null) {
			df = new SimpleDateFormat(ISO_DATE_FORMAT);
			isoDateThread.set(df);
		}

		return df;
	}
	
	/**
	 * DateFormat for parsing date like: yyyy-MM-dd-HH:mm:ss
	 * @return
	 */
	public static DateFormat getISODateTimeFormat() {
		DateFormat df = isoDateTimeThread.get();
		if (df == null) {
			df = new SimpleDateFormat(ISO_DATETIME_FORMAT);
			isoDateTimeThread.set(df);
		}

		return df;
	}
	
	// date format: yyyy-MM-dd'T'HH:mm:ss.SSS'Z'
	public static DateFormat getISODateTimeZoneFormat() {
		DateFormat df = isoDateTimeZoneThread.get();
		if (df == null) {
			df = new SimpleDateFormat(UTC_DATETIME_FORMAT);
			isoDateTimeZoneThread.set(df);
		}

		return df;
	}

	/**
	 * Parse date format: yyyy.MM.dd
	 * @param dateStr
	 * @param zone
	 * @return
	 * @throws ParseException
	 */
	public static Date parseDotDate(String dateStr, TimeZone zone) throws ParseException {
		DateFormat df = getDotDateFormat();
		if (zone != null) df.setTimeZone(zone);
		
		return df.parse(dateStr);
	}
	
	/**
	 * Parse date format: yyyy-MM-dd
	 * @param dateStr
	 * @param zone
	 * @return
	 * @throws ParseException
	 */
	public static Date parseISODate(String dateStr, TimeZone zone) throws ParseException {
		DateFormat df = getISODateFormat();
		if (zone != null) df.setTimeZone(zone);
		
		return df.parse(dateStr);
	}
	
	/**
	 * Date format: yyyy-MM-dd-HH:mm:ss
	 * @param dateStr
	 * @param zone
	 * @return
	 * @throws ParseException
	 */
	public static Date parseISODateTime(String dateStr, TimeZone zone) throws ParseException {
		DateFormat df = getISODateTimeFormat();
		if (zone != null) df.setTimeZone(zone);
		
		return df.parse(dateStr);
	}
	
	/**
	 * Parse date format: yyyy-MM-dd'T'HH:mm:ss.SSS'Z'
	 * @param dateStr
	 * @param zone
	 * @return
	 * @throws ParseException
	 */
	public static Date parseUTCDateTime(String dateStr, TimeZone zone) throws ParseException {
		DateFormat format = getISODateTimeZoneFormat();
		format.setTimeZone(zone);
		return format.parse(dateStr);
	}
	
	/**
	 * Format date to : yyyy.MM.dd
	 * @param date
	 * @param zone
	 * @return
	 */
	public static String formatTime(Date date) {
		DateFormat df = getTimeFormat();
		
		return df.format(date);
	}

	/**
	 * Format date to : yyyy.MM.dd
	 * @param date
	 * @param zone
	 * @return
	 */
	public static String formatDotDate(Date date, TimeZone zone) {
		DateFormat df = getDotDateFormat();
		if (zone != null) df.setTimeZone(zone);
		
		return df.format(date);
	}

	/**
	 * Format date to : yyyy-MM-dd
	 * @param date
	 * @param zone
	 * @return
	 */
	public static String formatISODate(Date date, TimeZone zone) {
		DateFormat df = getISODateFormat();
		if (zone != null) df.setTimeZone(zone);
		
		return df.format(date);
	}

	/**
	 * Format date time to : yyyy-MM-dd HH:mm:ss
	 * @param date
	 * @param zone
	 * @return
	 */
	public static String formatISODateTime(Date date, TimeZone zone) {
		DateFormat df = getISODateTimeFormat();
		if (zone != null) df.setTimeZone(zone);
		
		return df.format(date);
	}
	
	/**
	 * Format date time to : yyyy-MM-dd'T'HH:mm:ss.SSS'Z'
	 * @param date
	 * @param zone
	 * @return
	 */
	public static String formatUTCDateTime(Date date, TimeZone zone) {
		DateFormat format = getISODateTimeZoneFormat();
		format.setTimeZone(zone);
		return format.format(date);
	}
	
	/**
	 * Date stored in database may be not system time, for example, it may be beijing time. If we compare it with system time directly, it's not correct.
	 * We must convert it into system time first. 
	 * @param date Date from database
	 * @param zone Which time zone date does belong.
	 * @return Local time
	 */
	public static Date convertToSystemTime(Date date, TimeZone zone) {
		long gap = zone.getRawOffset() - TimeZone.getDefault().getRawOffset();
		
		return new Date(date.getTime() + gap);
	}

	/**
	 * Get the beginning datetime of this week.
	 *
	 * @return
	 */
	public static Date getStartDateTimeOfThisWeek(Date today) {
		Calendar cal = Calendar.getInstance();
		if (today != null) {
			cal.setTime(today);
		}
		cal.set(Calendar.DAY_OF_WEEK, Calendar.SUNDAY);
		cal.set(Calendar.HOUR_OF_DAY, 0);
		cal.set(Calendar.MINUTE, 0);
		cal.set(Calendar.SECOND, 0);
		cal.set(Calendar.MILLISECOND, 0);
		return cal.getTime();
	}

	/**
	 * Get the end datetime of this week.
	 *
	 * @return
	 */
	public static Date getEndDateTimeOfThisWeek(Date today) {
		Calendar cal = Calendar.getInstance();
		if (today != null) {
			cal.setTime(today);
		}
		cal.set(Calendar.DAY_OF_WEEK, Calendar.SATURDAY);
		cal.set(Calendar.HOUR_OF_DAY, 23);
		cal.set(Calendar.MINUTE, 59);
		cal.set(Calendar.SECOND, 59);
		cal.set(Calendar.MILLISECOND, 999);
		return cal.getTime();
	}

	/**
	 * Get the date seven days ago.
	 *
	 * @return
	 */
	public static Date getSevenDaysAgoDate(Date today) {
		Calendar cal = Calendar.getInstance();
		if (today != null) {
			cal.setTime(today);
		}
		cal.set(Calendar.DAY_OF_YEAR, cal.get(Calendar.DAY_OF_YEAR) - 7);
		cal.set(Calendar.HOUR_OF_DAY, 0);
		cal.set(Calendar.MINUTE, 0);
		cal.set(Calendar.SECOND, 0);
		cal.set(Calendar.MILLISECOND, 0);
		return cal.getTime();
	}

	/**
	 * Parse a date string into a Date object.
	 * 
	 * @param text
	 *            Date string representation.
	 * @return Date Object
	 */
	public static Date resolveDate(String text) {
		Date date = null;
		try {
			date = getISODateFormat().parse(text);
		} catch (ParseException e) {
			try {
				date = getSlashDateFormat().parse(text);
			} catch (ParseException e2) {
				try {
					date = getDotDateFormat().parse(text);
				} catch (ParseException e3) {
					logger.log(Level.WARNING,
							"Failed to parse text into Date object. Inputted text is: "
									+ text);
				}
			}
		}

		return date;
	}

	/**
	 * Get a date's time.
	 * @param date
	 * @return
	 */
	public static Time resolveTime(Date date) {
		String text = getTimeFormat().format(date);
		return Time.valueOf(text);
	}
	
	/**
	 * Return the absolute time lag between two dates.
	 * @param date1
	 * @param date2
	 * @return absolute milliseconds between two dates.
	 */
	public static long timeLag(Date date1, Date date2) {
		return Math.abs(date1.getTime() - date2.getTime());
	}

}
