package com.info.reader;

import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public interface IExcelReader {

	public void closeFile();

	public Sheet getSheetByName(String sheetName);

	public Row getRowByRowNumber(String sheetName, int rowNumber);

	public List<String> getAllPresentColums(Row row);
	
	public List<String> getAllCellsByColumnName(String sheetName, int coulmnRowNumber, String columnName);

}
