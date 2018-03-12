package com.anlohse.porter;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Locale.Category;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.anlohse.porter.internal.RowMapping;

public class ExcelParser {
	
	Locale locale;
	boolean ignoreBlank;
	
	public ExcelParser() {
		this(Locale.getDefault(Category.FORMAT));
	}

	public ExcelParser(Locale locale) {
		this.locale = locale;
	}

	public <T> List<T> parseXls(InputStream is, Class<T> clazz, int indexOfSheet, boolean hasHeader) throws IOException, ExcelParserException {
		RowMapping<T> mapping = RowMapping.getMapping(clazz);
		ArrayList<T> list = new ArrayList<>();
		parseXsl(is, mapping, list, indexOfSheet, hasHeader);
		return list;
	}
	
	public <T> List<T> parseXls(InputStream is, Class<T> clazz) throws IOException, ExcelParserException {
		return parseXls(is, clazz, 0, false);
	}

	private <T> void parseXsl(InputStream is, RowMapping<T> mapping, ArrayList<T> list, int indexOfSheet, boolean hasHeader)
			throws IOException, ExcelParserException {
		HSSFWorkbook wb = new HSSFWorkbook(is);
		try {
			HSSFSheet sheet = wb.getSheetAt(indexOfSheet);
	
			parseASheet(mapping, list, sheet, hasHeader);
		} finally {
			wb.close();
		}
	}

	public <T> List<T> parseXlsx(InputStream is, Class<T> clazz, int indexOfSheet, boolean hasHeader) throws IOException, ExcelParserException {
		RowMapping<T> mapping = RowMapping.getMapping(clazz);
		ArrayList<T> list = new ArrayList<>();
		parseXslx(is, mapping, list, indexOfSheet, hasHeader);
		return list;
	}

	public <T> List<T> parseXlsx(InputStream is, Class<T> clazz) throws IOException, ExcelParserException {
		return parseXlsx(is, clazz, 0, false);
	}

	private <T> void parseXslx(InputStream is, RowMapping<T> mapping, ArrayList<T> list, int indexOfSheet, boolean hasHeader)
			throws IOException, ExcelParserException {
		XSSFWorkbook wb = new XSSFWorkbook(is);
		try {
			XSSFSheet sheet = wb.getSheetAt(indexOfSheet);
	
			parseASheet(mapping, list, sheet, hasHeader);
		} finally {
			wb.close();
		}
	}

	public Locale getLocale() {
		return locale;
	}

	public ExcelParser setLocale(Locale locale) {
		this.locale = locale;
		return this;
	}

	public List<Object[]> parseXls(InputStream is, int indexOfSheet, boolean hasHeader) throws IOException, ExcelParserException {
		ArrayList<Object[]> list = new ArrayList<>();
		parseXsl(is, list, indexOfSheet, hasHeader);
		return list;
	}
	
	public List<Object[]> parseXls(InputStream is) throws IOException, ExcelParserException {
		return parseXls(is, 0, false);
	}

	private void parseXsl(InputStream is, ArrayList<Object[]> list, int indexOfSheet, boolean hasHeader)
			throws IOException, ExcelParserException {
		HSSFWorkbook wb = new HSSFWorkbook(is);
		try {
			HSSFSheet sheet = wb.getSheetAt(indexOfSheet);
	
			parseASheet(null, list, sheet, hasHeader);
		} finally {
			wb.close();
		}
	}

	public List<Object[]> parseXlsx(InputStream is, int indexOfSheet, boolean hasHeader) throws IOException, ExcelParserException {
		ArrayList<Object[]> list = new ArrayList<>();
		parseXslx(is, list, indexOfSheet, hasHeader);
		return list;
	}

	public List<Object[]> parseXlsx(InputStream is) throws IOException, ExcelParserException {
		return parseXlsx(is, 0, false);
	}

	private void parseXslx(InputStream is, ArrayList<Object[]> list, int indexOfSheet, boolean hasHeader)
			throws IOException, ExcelParserException {
		XSSFWorkbook wb = new XSSFWorkbook(is);
		try {
			XSSFSheet sheet = wb.getSheetAt(indexOfSheet);
	
			parseASheet(null, list, sheet, hasHeader);
		} finally {
			wb.close();
		}
	}

	@SuppressWarnings("unchecked")
	private <T> void parseASheet(RowMapping<T> mapping, ArrayList<T> list, Sheet sheet, boolean hasHeader) throws ExcelParserException {
		Iterator<Row> rows = sheet.rowIterator();
		Object[] values = new Object[16]; 
		if (hasHeader && rows.hasNext())
			rows.next(); // jump first row
		while (rows.hasNext()) {
			Arrays.fill(values, 0, values.length, null);
			Row row = rows.next();
			Iterator<Cell> cells = row.cellIterator();
			int i = 0;
			boolean notblank = true;
			while (cells.hasNext()) {
				Cell cell = cells.next();
				notblank |= cell.getCellTypeEnum() != CellType.BLANK;
				i = cell.getColumnIndex();
				if (i > values.length)
					values = Arrays.copyOf(values, values.length * 2);
				switch (cell.getCellTypeEnum()) {
				case STRING:  values[i] = cell.getStringCellValue(); break;
				case NUMERIC: {
					if (DateUtil.isCellDateFormatted(cell))
						values[i] = cell.getDateCellValue(); 
					else
						values[i] = cell.getNumericCellValue(); 
				} break;
				case BOOLEAN: values[i] = cell.getBooleanCellValue(); break;
				case FORMULA: {
					switch (cell.getCachedFormulaResultTypeEnum()) {
					case STRING:  values[i] = cell.getStringCellValue(); break;
					case NUMERIC: {
						if (DateUtil.isCellDateFormatted(cell))
							values[i] = cell.getDateCellValue(); 
						else
							values[i] = cell.getNumericCellValue(); 
					} break;
					case BOOLEAN: values[i] = cell.getBooleanCellValue(); break;
					case BLANK:
					case ERROR:   values[i] = null; break;
					}
					break;
				}
				case BLANK:
				case ERROR:   values[i] = null; break;
				}
			}
			if (notblank || !ignoreBlank) {
				if (mapping != null)
					list.add(mapping.createRow(values,locale));
				else
					list.add((T)Arrays.copyOf(values, i + 1));
			}
		}
	}

	public boolean isIgnoreBlank() {
		return ignoreBlank;
	}

	public ExcelParser setIgnoreBlank(boolean ignoreBlank) {
		this.ignoreBlank = ignoreBlank;
		return this;
	}

}
