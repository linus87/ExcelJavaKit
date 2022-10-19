package com.linus.excel;

import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.linus.excel.validation.ColumnConstraint;
import com.linus.excel.validation.IntegerRangeColumnConstraint;
import com.linus.excel.validation.RangeColumnConstraint;

/**
 * 
 * @author lyan2
 */
public class MapSheetWriter extends AbstractSheetWriter<Map<String, Object>> {
    private final Logger logger = Logger.getLogger(MapSheetWriter.class.getName());
    
	public MapSheetWriter(Workbook book, List<ColumnConfiguration> configs) {
        super(book, configs);
    }

	@Override
	public void writeRow(Workbook book, Sheet sheet, Row row, Map<String, Object> map) {
		for (ColumnConfiguration config : configs) {
			if (config != null) {
				CellStyle cellStyle = this.getDataCellStyle(config.getColumnIndex());
				createCell(book, sheet, row, config, map.get(config.getKey()), cellStyle);
			}
		}
	}

	@Override
	public void writeSheet(Workbook book, Sheet sheet, List<Map<String, Object>> list, boolean hasTitle) {
	    
		if (hasTitle)
			createTitle(book, sheet, configs);

		int rowNum = firstDataRowNum;
		for (Map<String, Object> map : list) {
			Row row = sheet.createRow(rowNum++);
			writeRow(book, sheet, row, map);
		}

		for (ColumnConfiguration config : configs) {
			if (!config.getDisplay()) {
				hideColumn(sheet, config.getColumnIndex());
			}
		}

		if (firstDataRowNum > sheet.getLastRowNum()) {
			// no data
			return;
		}

		for (ColumnConfiguration config : configs) {
			List<ColumnConstraint> constraints = config.getConstraints();
			for (ColumnConstraint constraint : constraints) {
				if (constraint instanceof RangeColumnConstraint) {
					// only support single
					if (((RangeColumnConstraint) constraint).isAllowMultiple()) break;
					
					XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
					XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
					    dvHelper.createExplicitListConstraint(((RangeColumnConstraint) constraint).getPickList());

					CellRangeAddressList addressList = new CellRangeAddressList(firstDataRowNum,  sheet.getLastRowNum(), config.getColumnIndex(), config.getColumnIndex());
					XSSFDataValidation validation =(XSSFDataValidation)dvHelper.createValidation(dvConstraint, addressList);

					// Display pick list when user click the cell.
					validation.setSuppressDropDownArrow(true);

					// Note this extra method call. If this method call is
					// omitted, or if the
					// boolean value false is passed, then Excel will not
					// validate the value the
					// user enters into the cell.
					validation.setShowErrorBox(((RangeColumnConstraint) constraint).getMustInRange());
					sheet.addValidationData(validation);
					break;
				}

				if (constraint instanceof IntegerRangeColumnConstraint) {
					IntegerRangeColumnConstraint iconstraint = (IntegerRangeColumnConstraint) constraint;
					XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
					XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
							.createNumericConstraint(XSSFDataValidationConstraint.ValidationType.INTEGER,
									XSSFDataValidationConstraint.OperatorType.BETWEEN, "=" + iconstraint.getMin(), "="
											+ iconstraint.getMax());

					CellRangeAddressList addressList = new CellRangeAddressList(firstDataRowNum, sheet.getLastRowNum(),
							config.getColumnIndex(), config.getColumnIndex());
					XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint,
							addressList);

					// Display pick list when user click the cell.
					validation.setSuppressDropDownArrow(true);

					// Note this extra method call. If this method call is
					// omitted, or if the
					// boolean value false is passed, then Excel will not
					// validate the value the
					// user enters into the cell.
					validation.setShowErrorBox(true);
					sheet.addValidationData(validation);
					break;
				}
			}
		}

		adjustColumnWidth(sheet, configs);
	}
	
	@Override
	public void createTitle(Workbook book, Sheet sheet, List<ColumnConfiguration> configs) {
		Row row = sheet.createRow(firstDataRowNum++);
		CellStyle headerStyle = this.getHeaderCellStyle();
		
		for (ColumnConfiguration config : configs) {
			if (config != null) {
				createCell(book, sheet, row, config.getColumnIndex(), config.getTitle(), headerStyle);
			}
		}
	}

	public void createSubHead(Workbook book, Sheet sheet, List<ColumnConfiguration> configs) {
		Row row = sheet.createRow(firstDataRowNum++);
		CellStyle headerStyle = book.createCellStyle();
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setFillForegroundColor(IndexedColors.LIME.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		headerStyle.setFont(this.defaultFont);
		headerStyle.setWrapText(true);

		for (ColumnConfiguration config : configs) {
			if (config != null) {
				createCell(book, sheet, row, config.getColumnIndex(), config.getLabel(), headerStyle);
			}
		}
	}

	/**
	 * Create a row filled with sample data.
	 * 
	 * @param book
	 * @param sheet
	 * @param configs
	 */
	public void createSampleRow(Workbook book, Sheet sheet, List<ColumnConfiguration> configs) {
		Row row = sheet.createRow(firstDataRowNum++);

		// sample data is locked;
		CellStyle style = book.createCellStyle();
		style.setLocked(true);
		style.setWrapText(true);
		style.setFont(defaultFont);

		for (ColumnConfiguration config : configs) {
			if (config != null) {
				createCell(book, row, config, config.getSample(), style);
			}
		}
	}

}
