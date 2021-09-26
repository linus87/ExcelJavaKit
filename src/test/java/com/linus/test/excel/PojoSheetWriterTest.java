package com.linus.test.excel;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.ResourceBundle;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import com.linus.excel.ColumnConfiguration;
import com.linus.excel.PojoSheetWriter;
import com.linus.excel.po.User;
import com.linus.excel.util.ColumnConfigurationParserForPojo;

public class PojoSheetWriterTest {

	@Test
	public void pojoToExcelTest() throws IntrospectionException, IOException {
		PojoSheetWriter<User> writer = new PojoSheetWriter<User>();
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Detail");
		
		ArrayList<ColumnConfiguration> configs = ColumnConfigurationParserForPojo.getColumnConfigurations(
				User.class, Locale.SIMPLIFIED_CHINESE, ResourceBundle.getBundle("sheet_header", Locale.SIMPLIFIED_CHINESE));
		
		writer.writeSheet(wb, sheet, configs, getUserList(), false);
		
		File file = new File("excel/pojo/user.xlsx");
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
	}
	
	public List<User> getUserList() {
		List<User> list = new ArrayList<User>(1);
		
		User u = new User();
		u.setAge(18);
		u.setBirthday(new Date());
		u.setEmail("lyan@ebay.com");
		
		list.add(u);
		
		u.setAge(25);
		u.setBirthday(new Date());
		u.setEmail("lyan2@ebay.com");
		
		return list;
	}
}
