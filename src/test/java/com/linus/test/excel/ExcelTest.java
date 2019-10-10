package com.linus.test.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.validation.ConstraintViolation;

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
import com.linus.excel.util.ExcelUtil;
import com.linus.excel.validation.ColumnConstraint;
import com.linus.excel.validation.NotNullColumnConstraint;

import junit.framework.Assert;

public class ExcelTest {
	private final Logger logger = Logger.getLogger(ExcelTest.class.getName());
	
	private ObjectMapper mapper = new ObjectMapper();
	
//	@Test
	public void testReaderRowAsList() throws IOException {
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
		adjustColumnConfigurations(configs);
		
		logger.log(Level.INFO, mapper.writeValueAsString(configs));
		
		File file = new File("excel/Listing_Template_203433884.xlsx");
		FileInputStream fis = new FileInputStream(file);
		Workbook wb = new XSSFWorkbook(fis);
		
		Sheet sheet = wb.getSheetAt(0);
		
		List<List<Object>> list = sheetReader.readSheet2(sheet, configs, 3, constraintViolations);	
		logger.log(Level.INFO, mapper.writeValueAsString(list));
		
		if (constraintViolations.size() > 0) {
			for (ConstraintViolation<Object> error : constraintViolations) {
				logger.log(Level.INFO, error.getMessage() + error.getInvalidValue());
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
	
	@Test
	public void testReaderRowAsMapWithValidation() throws IOException {
		Set<ConstraintViolation<Object>> constraintViolations = new HashSet<ConstraintViolation<Object>>();
		ISheetReader sheetReader = new SheetReader();
		
		File configFile = new File(ExcelTest.class.getResource("deals.json").getFile());
		JsonNode tree = mapper.readTree(configFile);
		ArrayList<ColumnConfiguration> configs = ExcelUtil.getColumnConfigurations((ArrayNode)tree, Locale.CHINA);
		adjustColumnConfigurations(configs);
		
		logger.log(Level.INFO, mapper.writeValueAsString(configs));
		
		File file = new File("excel/template.xlsx");
		FileInputStream fis = new FileInputStream(file);
		Workbook wb = new XSSFWorkbook(fis);
		
		Sheet sheet = wb.getSheetAt(0);
		
		List<Map<String, Object>> list = sheetReader.readSheet(sheet, configs, 3, constraintViolations);	
		
		if (constraintViolations.size() > 0) {
			for (ConstraintViolation<Object> error : constraintViolations) {
				logger.log(Level.INFO, error.getMessage() + error.getInvalidValue());
			}
		}
		
		if (list != null && !list.isEmpty()) {
			Iterator<Map<String, Object>> iter = list.iterator();
			while (iter.hasNext()) {
				Map<String, Object> user = iter.next();
				logger.log(Level.INFO, mapper.writeValueAsString(user));
			}
		}
		
		Assert.assertTrue(list.size() > 0);
		fis.close();
		wb.close();
	}
	
	/**
	 * Listing fields of promotion doesn't contain nomination id. But we need nomination id to differentiate listings. So we have to 
	 * prepend nomination id as the first column configuration.
	 * @param columnConfigs
	 */
	public List<ColumnConfiguration> adjustColumnConfigurations(List<ColumnConfiguration> columnConfigs) {
		
		// it's used to configure a hidden column, it will store nomination id.
		ColumnConfiguration nominationConfig = new ColumnConfiguration();
		nominationConfig.setKey("skuId");
		nominationConfig.setColumnIndex(0);
		nominationConfig.setWritable(false);
		nominationConfig.setDisplay(false);
		nominationConfig.setRawType("string");
//		nominationConfig.setSample(promoId);
		
		ColumnConstraint constraint = new NotNullColumnConstraint();
		constraint.setMessage("excel.validation.template.message");
		nominationConfig.getConstraints().add(constraint);
		
/*		// upload config
		ColumnConfiguration uploadConfig = new ColumnConfiguration();
		uploadConfig.setKey("toUpload");
		uploadConfig.setTitle("excel.header.toUpload");
		//uploadConfig.setLabel("Whether Upload");
		uploadConfig.setSample("excel.header.uploadSample");
		uploadConfig.setReadOrder(1);
		uploadConfig.setWriteOrder(1);
		uploadConfig.setWritable(true);
		uploadConfig.setDisplay(true);
		uploadConfig.setRawType("picklist");
		
		RangeColumnConstraint rangeConstraint = new RangeColumnConstraint();
		String[] whetherToUpload = {"Y", "N"};
		rangeConstraint.setPickList(whetherToUpload);
		uploadConfig.getConstraints().add(rangeConstraint);*/
		
		if (columnConfigs != null) {
			
			// move colmuns to right by one column
			for (ColumnConfiguration config : columnConfigs) {
				config.setColumnIndex(config.getColumnIndex() + 1);
			}
			
			columnConfigs.add(nominationConfig);
//			columnConfigs.add(uploadConfig);
		}
		
		return columnConfigs;
	}
}
