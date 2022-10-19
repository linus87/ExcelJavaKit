package com.linus.test.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
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
    
    private void createOptions(Workbook wb, List<String> values, String optionName, String optionLabel) {
        Sheet sheet = wb.getSheet("options");
        if (sheet == null) {
            sheet = wb.createSheet("options");
        }
        
        Name namedArea = wb.createName();
        namedArea.setNameName(optionName);
        
        Row optionLabelRow = sheet.createRow(0);
        int columnIndex = optionLabelRow.getLastCellNum() + 1;
        optionLabelRow.createCell(columnIndex).setCellValue(optionLabel);
        
        int rowIndex = 1;
        
        for (String value : values) {
            Row row = sheet.createRow(rowIndex++);
            row.createCell(columnIndex).setCellValue(value);
        }
        
        String colStr = CellReference.convertNumToColString(columnIndex);
        String formular = String.format("%s!$%s$2:$%s$%d", sheet.getSheetName(), colStr, colStr, rowIndex);
        namedArea.setRefersToFormula(formular);
    }
    
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

        // write to excel
        sheetWriter.writeSheet(wb, sheet, list, true);
        
        createOptions(wb, Arrays.asList(options), "warehouse", "Warehouse");
        
        createDropdown(wb, sheet, 1, "warehouse");
        
        wb.write(fos);
        fos.close();
        wb.close();
    }
    
    private void createDropdown(Workbook wb, Sheet sheet, int columnIndex, String optionsName) {
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
        
        DataValidationConstraint constraint = dvHelper.createFormulaListConstraint(optionsName);
        
        CellRangeAddressList addressList = new CellRangeAddressList(1,  sheet.getLastRowNum(), columnIndex, columnIndex);
        DataValidation dv = dvHelper.createValidation(constraint, addressList);
        sheet.addValidationData(dv);
    }
}
