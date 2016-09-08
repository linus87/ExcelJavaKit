package com.linus.test.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.linus.excel.ColumnConfiguration;
import com.linus.excel.SheetWriter;
import com.linus.excel.util.ExcelUtil;
import com.linus.excel.validation.ColumnConstraint;
import com.linus.excel.validation.NotNullColumnConstraint;

public class ExcelWriterTest {
	private final Logger logger = Logger.getLogger(ExcelWriterTest.class.getName());
	
	private ObjectMapper mapper = new ObjectMapper();
	
	ResourceBundle bundle;
	
	@Before
	public void before() {
		bundle = ResourceBundle.getBundle("ExcelValidationMessages") ;
	}
	
	@Test
	public void testWriter() throws IOException {
		File file = new File("excel/template.xlsx");
		FileOutputStream fos = new FileOutputStream(file);
		Workbook wb = new XSSFWorkbook();
		
		SheetWriter sheetWriter = new SheetWriter();
		Sheet sheet = wb.createSheet();
		
		File configFile = new File(ExcelTest.class.getResource("deals.json").getFile());
		JsonNode tree = mapper.readTree(configFile);
		ArrayList<ColumnConfiguration> configs = ExcelUtil.getColumnConfigurations((ArrayNode)tree, Locale.CHINA);
		adjustColumnConfigurations(configs);
		
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		Map<String, Object> map = new HashMap<String, Object>();
		
		list = mapper.readValue("[{\"skuId\":\"a1qO0000001KVV3IAO\",\"skuName\":\"superman_2015 iKross 360掳 Car Air Vent Mount Cradle Holder Stand For Mobile Phone Cell Phone\",\"itemId\":\"6462738543254355738373746\",\"itemTitle\":null,\"currPrice\":{\"value\":34,\"currency\":\"USD\"},\"dealsPrice\":null,\"proposePrice\":null,\"stockNum\":null,\"stockReadyDate\":null,\"currency\":\"USD\",\"state\":null},{\"skuId\":\"a1qO0000001KVV2IAO\",\"skuName\":\"superman_2015 iKross 360掳 Car Air Vent Mount Cradle Holder Stand For Mobile Phone Cell Phone\",\"itemId\":null,\"itemTitle\":null,\"currPrice\":null,\"dealsPrice\":null,\"proposePrice\":null,\"stockNum\":null,\"stockReadyDate\":null,\"currency\":\"USD\",\"state\":null}]", List.class);
		preHandleData(configs, list);
//		Object[] values = {"hello world", new Date(), true, 34.8, 8};
//		LinkedHashMap<String, Object> price = new LinkedHashMap<String, Object>();
//		price.put("value", 34);
//		price.put("currency", "USD");
//		Object[] values = {"hello world", price, price, 1111111111111111l, price, price, price, price};
		
//		if (configs.size() > 0) {
//			int index = 0;
//			for (ColumnConfiguration config : configs) {
//				map.put(config.getKey(), values[index++ % 5]);
//			}
//		}
		
//		list.add(map);
//		list.add(map);
		sheetWriter.writeSheet(wb, sheet, configs, list, true);
		sheetWriter.freeze(sheet, 0, sheetWriter.getFirstDataRowNum());
		sheetWriter.setProtectionPassword(sheet, "111111");
		
		wb.write(fos);
		fos.close();
		wb.close();
	}
	
	private void preHandleData(List<ColumnConfiguration> columnConfigs, List<Map<String, Object>> listings) {
		if (listings != null && columnConfigs != null) {
			for (ColumnConfiguration config : columnConfigs) {
				if ("attachment".equalsIgnoreCase(config.getRawType())) {
					for(Map<String, Object> map : listings) {
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
	 * @param columnConfigs
	 */
	private List<ColumnConfiguration> adjustColumnConfigurations(List<ColumnConfiguration> columnConfigs) {
		// it's used to configure a hidden column, it will store nomination id.
		ColumnConfiguration nominationConfig = new ColumnConfiguration();
		nominationConfig.setKey("skuId");
		nominationConfig.setReadOrder(0);
		nominationConfig.setWriteOrder(0);
		nominationConfig.setWritable(false);
		nominationConfig.setDisplay(false);
		nominationConfig.setRawType("string");
		
		ColumnConstraint constraint = new NotNullColumnConstraint();
		constraint.setMessage("excel.validation.template.message");
		
		if (columnConfigs != null) {
			// move colmuns to right by one column
			for (ColumnConfiguration config : columnConfigs) {
				config.setReadOrder(config.getReadOrder() + 1);
				config.setWriteOrder(config.getWriteOrder() + 1);
			}
			
			columnConfigs.add(nominationConfig);
		}
		
		return columnConfigs;
	}
}
