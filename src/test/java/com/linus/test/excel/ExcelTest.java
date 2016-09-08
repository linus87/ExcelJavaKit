package com.linus.test.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;

import junit.framework.Assert;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.linus.excel.ColumnConfiguration;
import com.linus.excel.ISheetReader;
import com.linus.excel.SheetReader;
import com.linus.excel.po.User;
import com.linus.excel.util.ExcelUtil;

public class ExcelTest {
	private final Logger logger = Logger.getLogger(ExcelTest.class.getName());
	
	private ObjectMapper mapper = new ObjectMapper();
	
	@Test
	public void testReaderRowAsList() throws IOException {
		Set<ConstraintViolation<Object>> constraintViolations = new HashSet<ConstraintViolation<Object>>();
		ISheetReader sheetReader = new SheetReader();
		
		File file = new File("excel/template.xlsx");
		FileInputStream fis = new FileInputStream(file);
		Workbook wb = new XSSFWorkbook(fis);
		
		Sheet sheet = wb.getSheetAt(0);
		List<List<Object>> list = sheetReader.readSheet(sheet, 3);	
		
		if (list != null && !list.isEmpty()) {
			Iterator<List<Object>> iter = list.iterator();
			while (iter.hasNext()) {
				List<Object> user = iter.next();
				logger.log(Level.INFO, mapper.writeValueAsString(user));
			}
		}
		
		Assert.assertTrue(list.size() > 0);
		fis.close();
		wb.close();
	}
	
//	@Test
	public void testReaderRowAsListWithValidation() throws IOException {
		Set<ConstraintViolation<Object>> constraintViolations = new HashSet<ConstraintViolation<Object>>();
		ISheetReader sheetReader = new SheetReader();
		
		File configFile = new File(ExcelTest.class.getResource("deals.json").getFile());
		JsonNode tree = mapper.readTree(configFile);
		ArrayList<ColumnConfiguration> configs = ExcelUtil.getColumnConfigurations((ArrayNode)tree, Locale.CHINA);
		
		if (configs.size() > 0) {
			for (ColumnConfiguration config : configs) {
				logger.log(Level.INFO, config.getKey());
			}
		}
		
		File file = new File("excel/Listing_Template_203433884.xlsx");
		FileInputStream fis = new FileInputStream(file);
		Workbook wb = new XSSFWorkbook(fis);
		
		Sheet sheet = wb.getSheetAt(0);
		List<List<Object>> list = sheetReader.readSheet2(sheet, configs, 3, constraintViolations);	
		
		if (constraintViolations.size() > 0) {
			for (ConstraintViolation error : constraintViolations) {
				logger.log(Level.INFO, error.getMessage());
			}
		}
		
		if (list != null && !list.isEmpty()) {
			Iterator<List<Object>> iter = list.iterator();
			while (iter.hasNext()) {
				List<Object> user = iter.next();
				logger.log(Level.INFO, mapper.writeValueAsString(user));
			}
		}
		
		Assert.assertTrue(list.size() > 0);
		fis.close();
		wb.close();
	}
	
//	@Test
	public void testReader() throws IOException {
		Set<ConstraintViolation<Object>> constraintViolations = new HashSet<ConstraintViolation<Object>>();
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
		List<Object> users = sheetReader.readSheet(sheet, User.class, 1, constraintViolations);
		
		Assert.assertNotNull(constraintViolations);
		
		if (constraintViolations != null) {
			System.out.println(constraintViolations.size());
			Iterator<ConstraintViolation<Object>> violationIter = constraintViolations.iterator();
			while(violationIter.hasNext()) {
				ConstraintViolation<Object> error = violationIter.next();
				logger.log(Level.INFO, "Error message: " + error.getMessage());
				logger.log(Level.INFO, "Invalid: " + error.getInvalidValue());
			}
		}		
		
		if (users != null && !users.isEmpty()) {
			Iterator<Object> iter = users.iterator();
			while (iter.hasNext()) {
				User user = (User)iter.next();
				logger.log(Level.INFO, mapper.writeValueAsString(user));
			}
		}
		
		Assert.assertTrue(users.size() > 0);
		fis.close();
		wb.close();
	}
}
