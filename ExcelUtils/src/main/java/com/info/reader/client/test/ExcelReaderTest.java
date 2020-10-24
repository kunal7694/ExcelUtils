package com.info.reader.client.test;

import java.util.List;

import org.apache.poi.ss.usermodel.Row;

import com.info.reader.ExcelReader;

public class ExcelReaderTest {

	private static final String SAMPLE_XLSX_FILE_XLSX = System.getProperty("user.dir")
			+ "\\src\\main\\resources\\sample-xlsx-file.xlsx";

	public static void main(String[] args) {

		ExcelReader excelReader = new ExcelReader(SAMPLE_XLSX_FILE_XLSX);
		String sheetName = "Employee";
		
		Row row = excelReader.getRowByRowNumber(sheetName, 0);
		List<String> allPresentColums = excelReader.getAllPresentColums(row);
		System.out.println(allPresentColums);
		
		List<String> allCellsByColumnName = excelReader.getAllCellsByColumnName(sheetName, 0, "Name");
		System.out.println(allCellsByColumnName);

		excelReader.closeFile();

	}

}
