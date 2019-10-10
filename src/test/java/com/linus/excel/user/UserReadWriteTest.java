package com.linus.excel.user;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Time;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.linus.excel.ColumnConfiguration;
import com.linus.excel.ISheetReader;
import com.linus.excel.ISheetWriter;
import com.linus.excel.MapSheetWriter;
import com.linus.excel.SheetReader;
import com.linus.excel.enums.Gender;
import com.linus.excel.po.User;
import com.linus.excel.po.util.ExcelUtil;

import junit.framework.Assert;

public class UserReadWriteTest {
	private final Logger logger = Logger.getLogger(UserReadWriteTest.class.getName());
	private ObjectMapper mapper = new ObjectMapper();
	private ResourceBundle bundle;
	
	@Before
	public void init() {
		bundle = ResourceBundle.getBundle("sheet_header");
	}
	
	@Test
	public void testReader() throws IOException {
		Set<ConstraintViolation<User>> constraintViolations = new HashSet<ConstraintViolation<User>>();
		// preparing validation
//		ValidatorFactory factory = Validation.byDefaultProvider().configure().messageInterpolator(new ResourceBundleMessageInterpolator(new PlatformResourceBundleLocator("ExcelValidationMessages"))).buildValidatorFactory();
		ValidatorFactory factory = Validation.buildDefaultValidatorFactory();
		Validator validator = factory.getValidator();
		
		ISheetReader sheetReader = new SheetReader();
		sheetReader.setValidator(validator);
		
		File file = new File("excel/sheetreader.xlsx");
		FileInputStream fis = new FileInputStream(file);
		Workbook wb = new XSSFWorkbook(fis);
		
		Sheet sheet = wb.getSheetAt(0);
		List<User> users = sheetReader.readSheet(sheet, User.class, 1, constraintViolations);
		
		Assert.assertNotNull(constraintViolations);
		
		if (constraintViolations != null) {
			System.out.println(constraintViolations.size());
			Iterator<ConstraintViolation<User>> violationIter = constraintViolations.iterator();
			while(violationIter.hasNext()) {
				ConstraintViolation<User> error = violationIter.next();
				logger.log(Level.INFO, "Error message: " + error.getMessage());
				logger.log(Level.INFO, "Invalid: " + error.getInvalidValue());
			}
		}		
		
		if (users != null && !users.isEmpty()) {
			Iterator<User> iter = users.iterator();
			while (iter.hasNext()) {
				User user = (User)iter.next();
				logger.log(Level.INFO, mapper.writeValueAsString(user));
			}
		}
		
		Assert.assertTrue(users.size() > 0);
		fis.close();
		wb.close();
	}
	
	@Test
	public void testWriter() throws IOException, IntrospectionException {
		ISheetWriter<Map<String, Object>> sheetWriter = new MapSheetWriter();
		
		Workbook wb = new XSSFWorkbook();
		
		Sheet sheet = wb.createSheet();
		
		ArrayList<ColumnConfiguration> configs = ExcelUtil.getColumnConfigurations(User.class, Locale.CHINA, bundle);
		
		List<Map<String, Object>> users = new ArrayList<Map<String, Object>>(3);
		users.add(mapper.convertValue(createUser("Linus", "Yan", 30, Gender.MALE, "lyan2@ebay.com", BigDecimal.ONE, "yes", true, new Date(), "error", new Time(Calendar.getInstance().getTimeInMillis()), 0.1, Calendar.getInstance()), Map.class));
		users.add(mapper.convertValue(createUser("Linus", "Yan", 30, Gender.MALE, "lyan2@ebay.com", BigDecimal.ONE, "yes", true, new Date(), "error", new Time(Calendar.getInstance().getTimeInMillis()), 0.1, Calendar.getInstance()), Map.class));
		users.add(mapper.convertValue(createUser("Linus", "Yan", 30, Gender.MALE, "lyan2@ebay.com", BigDecimal.ONE, "yes", true, new Date(), "error", new Time(Calendar.getInstance().getTimeInMillis()), 0.1, Calendar.getInstance()), Map.class));
		
		sheetWriter.writeSheet(wb, sheet, configs, users, true);
		
		File file = new File("excel/user_writer.xlsx");
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
		
		fos.close();
		wb.close();
	}
	
	private User createUser(String firstName, String lastName, Integer age, Gender gender, String email, BigDecimal balance, String free, Boolean student, Date birthday, String error, Time time, Double completed, Calendar end) {
		User user = new User();
		user.setFirstName(firstName);
		user.setLastName(lastName);
		user.setAge(age);
		user.setGender(gender);
		user.setBirthday(birthday);
		user.setCompleted(0.1);
		user.setEmail(email);
		user.setEnd(end);
		user.setFree(free);
		user.setTime(time);
		user.setStudent(false);
		user.setBalance(balance);
		
		return user;
	}
}
