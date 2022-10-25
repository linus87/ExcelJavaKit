package com.linus.test.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.linus.excel.ColumnConfiguration;
import com.linus.excel.MapSheetWriter;
import com.linus.excel.util.ColumnConfigurationParserForJson;

public class DropdownListTest {

    private ObjectMapper mapper = new ObjectMapper();
    
    private String[] options = {"Hello", "World"};
    
    @SuppressWarnings("unchecked")
    @Test
    public void testWriter() throws IOException {
        File file = new File("excel/dropdown.xlsx");
        FileOutputStream fos = new FileOutputStream(file);
        Workbook wb = new XSSFWorkbook();
        
        Sheet sheet = wb.createSheet();

        // read configuration
        File configFile = new File(ExcelTest.class.getResource("config/configuration.json").getFile());
        JsonNode tree = mapper.readTree(configFile);
        ArrayList<ColumnConfiguration> configs = ColumnConfigurationParserForJson
                .getColumnConfigurations((ArrayNode) tree, Locale.CHINA);

        // read data
        List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
        File dataFile = new File(ExcelTest.class.getResource("data/data.json").getFile());
        JsonNode dataJson = mapper.readTree(dataFile);
        System.out.println(dataJson.toString());

        list = mapper.readValue(dataJson.toString(), List.class);

        MapSheetWriter sheetWriter = new MapSheetWriter(wb, configs);
        
        sheetWriter.createOptions(Arrays.asList(options), "warehouse");
        
        // write to excel
        sheetWriter.writeSheet(wb, sheet, list, true);
        
        sheetWriter.createDropdown(wb, sheet, 1, "warehouse");
        
        wb.write(fos);
        fos.close();
        wb.close();
    }
    
}
