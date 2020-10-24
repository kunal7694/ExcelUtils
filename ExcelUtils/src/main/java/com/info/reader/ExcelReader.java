package com.info.reader;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader implements IExcelReader {

	final private String filePath;

	private Workbook workbook = null;

	final Logger LOGGER = Logger.getLogger(ExcelReader.class.getName());

	private DataFormatter dataFormatter = new DataFormatter();

	private File file = null;

	public ExcelReader(String filePath) {
		this.filePath = filePath;
		this.loadFile();
	}

	private void loadFile() {
		file = new File(filePath);
		if (file.exists()) {
			try {
				String fileName = file.getName();
				LOGGER.log(Level.INFO, String.format("Started loding %s file", fileName));
				if (StringUtils.isNoneBlank(filePath)) {
					if (filePath.endsWith(".xlsx")) {
						workbook = new XSSFWorkbook(file);
						LOGGER.log(Level.INFO, fileName + " " + "File Loaded Succesfully");
					}
				}
			} catch (FileNotFoundException e) {
				LOGGER.log(Level.WARNING, String.format("File Not Found At %s Provided Location", filePath));
			} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
				LOGGER.log(Level.WARNING, e.getMessage() + "occured while creating workbook");
			}
		}
	}

	@Override
	public void closeFile() {
		if (workbook != null) {
			try {
				workbook.close();
				LOGGER.log(Level.INFO, String.format("File %s Closed sucessfully", file.getName()));
			} catch (IOException e) {
				LOGGER.log(Level.WARNING, e.getMessage() + "occured while closing workbook");
			}
		}
	}

	@Override
	public Sheet getSheetByName(String sheetName) {
		Sheet sheet = null;
		try {
			sheet = workbook.getSheet(sheetName);
		} catch (Exception e) {
			LOGGER.log(Level.WARNING, e.getMessage());
		}
		return sheet;
	}

	@Override
	public Row getRowByRowNumber(String sheetName, int rowNumber) {
		Row row = null;
		try {
			row = this.getSheetByName(sheetName).getRow(rowNumber);
		} catch (Exception e) {
			LOGGER.log(Level.WARNING, e.getMessage());
		}
		return row;
	}

	@Override
	public List<String> getAllPresentColums(Row row) {
		return extractRowData(row);
	}

	/**
	 * @param row
	 * @return
	 */
	private List<String> extractRowData(Row row) {
		Iterator<Cell> cellIterator = row.cellIterator();
		List<String> cells = new ArrayList<>();
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			cells.add(dataFormatter.formatCellValue(cell));
		}
		return cells;
	}

	@Override
	public List<String> getAllCellsByColumnName(String sheetName, int coulmnRowNumber, String columnName) {
		List<String> cells = new ArrayList<>();

		Row row = this.getRowByRowNumber(sheetName, coulmnRowNumber);
		List<String> allPresentColums = getAllPresentColums(row);

		try {
			int index = allPresentColums.indexOf(columnName);
			this.getSheetByName(sheetName).rowIterator().forEachRemaining(rows -> {
				cells.add(dataFormatter.formatCellValue(rows.getCell(index)));
			});
			// Removing column name form list
			cells.remove(0);
		} catch (Exception e) {
			LOGGER.log(Level.WARNING,
					String.format("Invalid %s column found, Please provide valid column name..", columnName));
		}
		return cells;
	}

}
