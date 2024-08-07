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
import com.linus.excel.ISheetWriter;
import com.linus.excel.InvalidRowError;
import com.linus.excel.MapSheetWriter;
import com.linus.excel.PojoSheetReader;
import com.linus.excel.enums.Gender;
import com.linus.excel.po.User;
import com.linus.excel.util.ColumnConfigurationParserForPojo;

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
		Set<InvalidRowError<User>> constraintViolations = new HashSet<InvalidRowError<User>>();
		// preparing validation
//		ValidatorFactory factory = Validation.byDefaultProvider().configure().messageInterpolator(new ResourceBundleMessageInterpolator(new PlatformResourceBundleLocator("ExcelValidationMessages"))).buildValidatorFactory();
		ValidatorFactory factory = Validation.buildDefaultValidatorFactory();
		Validator validator = factory.getValidator();
		
		PojoSheetReader<User> sheetReader = new PojoSheetReader<User>();
		sheetReader.setValidator(validator);
		
		File file = new File("excel/user_reader.xlsx");
		FileInputStream fis = new FileInputStream(file);
		Workbook wb = new XSSFWorkbook(fis);
		
		Sheet sheet = wb.getSheetAt(0);
		List<User> users = sheetReader.readSheet(sheet, User.class, 1, constraintViolations);
		
//		Assert.assertNotNull(constraintViolations);
		
		if (constraintViolations != null) {
			System.out.println(constraintViolations.size());
			Iterator<InvalidRowError<User>> violationIter = constraintViolations.iterator();
			while(violationIter.hasNext()) {
				InvalidRowError<User> error = violationIter.next();
				logger.log(Level.INFO, "Wrong property: " + error.getPropertyPath() + ", Error message: " + error.getMessage());
				logger.log(Level.INFO, "Invalid: " + mapper.writeValueAsString(error.getInvalidValue()));
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
		
		
		Workbook wb = new XSSFWorkbook();
		
		Sheet sheet = wb.createSheet();
		
		ArrayList<ColumnConfiguration> configs = ColumnConfigurationParserForPojo.getColumnConfigurations(User.class, Locale.CHINA, bundle);
		
		ISheetWriter<Map<String, Object>> sheetWriter = new MapSheetWriter(wb, configs);
		List<Map<String, Object>> users = new ArrayList<Map<String, Object>>(3);
		users.add(mapper.convertValue(createUser("Linus", "Yan", 30, Gender.MALE, "linus.yan@hotmail.com", BigDecimal.ONE, "yes", true, new Date(), "error", new Time(Calendar.getInstance().getTimeInMillis()), 0.1, Calendar.getInstance()), Map.class));
		users.add(mapper.convertValue(createUser("Linus", "Yan", 30, Gender.MALE, "linus.yan@hotmail.com", BigDecimal.ONE, "yes", true, new Date(), "error", new Time(Calendar.getInstance().getTimeInMillis()), 0.1, Calendar.getInstance()), Map.class));
		users.add(mapper.convertValue(createUser("Linus", "Yan", 30, Gender.MALE, "linus.yan@hotmail.com", BigDecimal.ONE, "yes", true, new Date(), "error", new Time(Calendar.getInstance().getTimeInMillis()), 0.1, Calendar.getInstance()), Map.class));
		
		sheetWriter.writeSheet(wb, sheet, users, true);
		
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
