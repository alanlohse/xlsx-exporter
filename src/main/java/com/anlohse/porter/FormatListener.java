package com.anlohse.porter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public interface FormatListener {

	void formatHeader(Cell cell, String columnName, int columnIndex, int rowIndex, String header);

	void formatCell(Cell cell, String columnName, int columnIndex, int rowIndex, Object value);

	void formatRow(Row row);

	void start(Workbook wb, Sheet sheet);

	void finish(Workbook wb, Sheet sheet);

}
