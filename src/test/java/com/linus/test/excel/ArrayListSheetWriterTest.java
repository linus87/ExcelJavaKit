package com.linus.test.excel;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import com.linus.excel.ArrayListSheetWriter;
import com.linus.excel.ColumnConfiguration;

public class ArrayListSheetWriterTest {

	@Test
	public void pojoToExcelTest() throws IntrospectionException, IOException {
		
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Detail");
		
		ArrayList<ColumnConfiguration> configs = getColumnConfigs();
		
		ArrayListSheetWriter<String> writer = new ArrayListSheetWriter<String>(wb, configs);
		writer.writeSheet(wb, sheet, getUserList(), true);
		
		File file = new File("excel/test/user.xlsx");
		if (!file.exists()) {
			file.createNewFile();
		}
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
	}
	
	public ArrayList<ColumnConfiguration> getColumnConfigs() {
		ArrayList<ColumnConfiguration> configs = new ArrayList<ColumnConfiguration>();
		
		ColumnConfiguration column1 = new ColumnConfiguration();
		column1.setColumnIndex(0);
		column1.setTitle("年龄");
		column1.setRawType("INTEGER");
		
		ColumnConfiguration column2 = new ColumnConfiguration();
		column2.setColumnIndex(1);
		column2.setTitle("姓名");
		column2.setRawType("STRING");
		
		ColumnConfiguration column3 = new ColumnConfiguration();
		column3.setColumnIndex(2);
		column3.setTitle("邮箱");
		column3.setRawType("STRING");
		
		configs.add(column1);
		configs.add(column2);
		configs.add(column3);
		
		return configs;
	}
	
	public List<String[]> getUserList() {
		List<String[]> list = new ArrayList<String[]>(2);
		
		String[] array1 = {"18", "test", "layn2@ebay.com"};
		String[] array2 = {"18", "test", "layn2@ebay.com"};
		
		list.add(array1);
		list.add(array2);
		return list;
	}
}
