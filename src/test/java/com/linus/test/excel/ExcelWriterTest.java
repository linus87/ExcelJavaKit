package com.linus.test.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.linus.excel.ColumnConfiguration;
import com.linus.excel.MapSheetWriter;
import com.linus.excel.util.ColumnConfigurationParserForJson;
import com.linus.excel.validation.ColumnConstraint;
import com.linus.excel.validation.NotNullColumnConstraint;

import javafx.scene.text.FontWeight;

public class ExcelWriterTest {
	private final Logger logger = Logger.getLogger(ExcelWriterTest.class.getName());

	private ObjectMapper mapper = new ObjectMapper();

	ResourceBundle bundle;

	@Before
	public void before() {
		bundle = ResourceBundle.getBundle("ExcelValidationMessages");
	}

	@SuppressWarnings("unchecked")
	@Test
	public void testWriter() throws IOException {
		File file = new File("excel/template.xlsx");
		FileOutputStream fos = new FileOutputStream(file);
		Workbook wb = new XSSFWorkbook();

		MapSheetWriter sheetWriter = new MapSheetWriter();
		Sheet sheet = wb.createSheet();

		// read configuration
		File configFile = new File(ExcelTest.class.getResource("config/configuration.json").getFile());
		JsonNode tree = mapper.readTree(configFile);
		ArrayList<ColumnConfiguration> configs = ColumnConfigurationParserForJson.getColumnConfigurations((ArrayNode) tree, Locale.CHINA);
		adjustColumnConfigurations(configs);

		// read data
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		File dataFile = new File(ExcelTest.class.getResource("data/data.json").getFile());
		JsonNode dataJson = mapper.readTree(dataFile);
		System.out.println(dataJson.toString());

		list = mapper.readValue(dataJson.toString(), List.class);

		// adjust configuration
		preHandleData(configs, list);

		// set font

		Font ft = wb.createFont();
		ft.setFontName("Arial");
		ft.setFontHeightInPoints((short) 9);
		sheetWriter.setDefaultFont(ft);
		
		Font titleFont = wb.createFont();
		titleFont.setFontName("Arial");
		titleFont.setFontHeightInPoints((short) 9);
		titleFont.setBold(true);
		sheetWriter.setTitleFont(titleFont);

		// write to excel
		sheetWriter.writeSheet(wb, sheet, configs, list, true);
		sheetWriter.freeze(sheet, 0, sheetWriter.getFirstDataRowNum());
		sheetWriter.setProtectionPassword(sheet, "123456");

		wb.write(fos);
		fos.close();
		wb.close();
	}

	private void preHandleData(List<ColumnConfiguration> columnConfigs, List<Map<String, Object>> listings) {
		if (listings != null && columnConfigs != null) {
			for (ColumnConfiguration config : columnConfigs) {
				if ("attachment".equalsIgnoreCase(config.getRawType())) {
					for (Map<String, Object> map : listings) {
						Object value = map.get(config.getKey());
						if (value == null) {
							map.put(config.getKey(), bundle.getString("listing.attachment.comment"));
						}
					}
				}
			}
		}
	}

	/**
	 * Nomination id is used to differentiate listings.
	 * 
	 * @param columnConfigs
	 */
	private List<ColumnConfiguration> adjustColumnConfigurations(List<ColumnConfiguration> columnConfigs) {
		// it's used to configure a hidden column, it will store nomination id.
		ColumnConfiguration nominationConfig = new ColumnConfiguration();
		nominationConfig.setKey("skuId");
		nominationConfig.setColumnIndex(0);
		nominationConfig.setWritable(false);
		nominationConfig.setDisplay(false);
		nominationConfig.setRawType("string");

		ColumnConstraint constraint = new NotNullColumnConstraint();
		constraint.setMessage("excel.validation.template.message");

		if (columnConfigs != null) {
			// move colmuns to right by one column
			for (ColumnConfiguration config : columnConfigs) {
				config.setColumnIndex(config.getColumnIndex() + 1);
			}

			columnConfigs.add(nominationConfig);
		}

		return columnConfigs;
	}
}
