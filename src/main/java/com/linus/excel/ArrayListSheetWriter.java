package com.linus.excel;

import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
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

public class ArrayListSheetWriter<T> extends AbstractSheetWriter<T[]> {

	@Override
	public void writeRow(Workbook book, Sheet sheet, Row row, List<ColumnConfiguration> configs, T[] array) {
		for (ColumnConfiguration config : configs) {
			if (config != null) {
				CellStyle cellStyle = getDataCellStyle(book, config.getColumnIndex());
				createCell(book, sheet, row, config, array[config.getColumnIndex()], cellStyle);
			}
		}
	}

	@Override
	public void writeSheet(Workbook book, Sheet sheet, List<ColumnConfiguration> configs, List<T[]> list,
			boolean hasTitle) {
		if (hasTitle)
			createTitle(book, sheet, configs);

		int rowNum = firstDataRowNum;
		for (T[] data : list) {
			Row row = sheet.createRow(rowNum++);
			writeRow(book, sheet, row, configs, data);
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

}
