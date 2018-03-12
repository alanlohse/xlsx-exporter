package com.anlohse.porter;

import java.io.IOException;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.Locale.Category;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.anlohse.porter.internal.RowMapping;
import com.anlohse.porter.types.Formula;

public class ExcelWriter {

	Locale locale;
	FormatListener formatListener;

	public ExcelWriter() {
		this(Locale.getDefault(Category.FORMAT));
	}

	public ExcelWriter(Locale locale) {
		this.locale = locale;
	}

	public <T> void writeXls(OutputStream out, List<T> objects, Class<T> clazz, String sheetName, boolean writeHeader) throws ExcelParserException, IOException {

		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = sheetName != null ? wb.createSheet(sheetName) : wb.createSheet();

		if (formatListener != null)
			formatListener.start(wb,sheet);

		writeSheet(objects, clazz, wb, sheet, writeHeader);

		if (formatListener != null)
			formatListener.finish(wb,sheet);

		wb.write(out);
		out.flush();
	}

	public <T> void writeXlsx(OutputStream out, List<T> objects, Class<T> clazz, String sheetName, boolean writeHeader) throws ExcelParserException, IOException {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = sheetName != null ? wb.createSheet(sheetName) : wb.createSheet();

		if (formatListener != null)
			formatListener.start(wb,sheet);

		writeSheet(objects, clazz, wb, sheet, writeHeader);

		if (formatListener != null)
			formatListener.finish(wb,sheet);

		wb.write(out);
		out.flush();
	}

	public void writeXls(OutputStream out, List<Object[]> objects, String sheetName) throws ExcelParserException, IOException {

		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = sheetName != null ? wb.createSheet(sheetName) : wb.createSheet();

		if (formatListener != null)
			formatListener.start(wb,sheet);

		writeSheet(objects, null, wb, sheet, false);

		if (formatListener != null)
			formatListener.finish(wb,sheet);

		wb.write(out);
		out.flush();
	}

	public void writeXlsx(OutputStream out, List<Object[]> objects, String sheetName) throws ExcelParserException, IOException {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = sheetName != null ? wb.createSheet(sheetName) : wb.createSheet();

		if (formatListener != null)
			formatListener.start(wb,sheet);

		writeSheet(objects, null, wb, sheet, false);

		if (formatListener != null)
			formatListener.finish(wb,sheet);

		wb.write(out);
		out.flush();
	}

	private <T> void writeSheet(List<T> objects, Class<T> clazz, Workbook wb, Sheet sheet, boolean writeHeader) throws ExcelParserException {
		RowMapping<T> mapping = clazz != null ? RowMapping.getMapping(clazz) : null;
		if (writeHeader) {
			Row row = sheet.createRow(0);
			if (formatListener != null)
				formatListener.formatRow(row);
			int n = 0;
			for (int c = 0; c < mapping.getCount(); c++) {
				if (mapping.getIndex(c) == n)
					n++;
				Cell cell = row.createCell(mapping.getIndex(c) != -1 ? mapping.getIndex(c) : n++);
				if (formatListener != null)
					formatListener.formatHeader(cell,mapping.getColumnName(c), c, 0, mapping.getHeader(c));
				cell.setCellValue(mapping.getHeader(c).toString());
			}
		}
		for (int r = 0; r < objects.size(); r++) {
			T rowObj = objects.get(r);
			Row row = sheet.createRow(r + (writeHeader ? 1 : 0));
			if (formatListener != null)
				formatListener.formatRow(row);
			Object[] values;
			if (mapping != null) {
				try {
					values = mapping.getValues(rowObj);
				} catch (Exception e) {
					throw new ExcelParserException(e);
				}
			} else {
				values = (Object[]) objects.get(r);
			}
			int n = 0;
			int cols = mapping != null ? mapping.getCount() : values.length;
			for (int c = 0; c < cols; c++) {
				if (mapping != null) {
					if (mapping.getIndex(c) == n)
						n++;
				}
				Cell cell = row.createCell(mapping != null ? (mapping.getIndex(c) != -1 ? mapping.getIndex(c) : n++) : c);
				if (formatListener != null)
					formatListener.formatCell(cell, mapping != null ? mapping.getColumnName(c) : null, c, r + (writeHeader ? 1 : 0), values[c]);
				if (values[c] instanceof String)
					cell.setCellValue((String) values[c]);
				else if (values[c] instanceof Number)
					cell.setCellValue(((Number) values[c]).doubleValue());
				else if (values[c] instanceof Date) {
					CellStyle cs = wb.createCellStyle();
					CreationHelper ch = wb.getCreationHelper();
					cs.setDataFormat(
							ch.createDataFormat().getFormat(getDateFormat() 
									+ (hasTime((Date) values[c]) ? " hh:mm:ss" : "")));
					cell.setCellValue((Date) values[c]);
					cell.setCellStyle(cs);
				} else if (values[c] instanceof Boolean)
					cell.setCellValue((Boolean) values[c]);
				else if (values[c] instanceof Formula)
					cell.setCellFormula(((Formula) values[c]).getFormula());
				else if (values[c] == null)
					cell.setCellValue((String) null);
				else
					cell.setCellValue(values[c].toString());
			}
		}
	}

	private String getDateFormat() throws ExcelParserException {
		SimpleDateFormat df = (SimpleDateFormat) DateFormat.getDateInstance(DateFormat.SHORT, locale);
		return df.toPattern();
	}

	private boolean hasTime(Date date) {
		Calendar cal = Calendar.getInstance(locale);
		return cal.get(Calendar.HOUR) != 0 && cal.get(Calendar.MINUTE) != 0 && cal.get(Calendar.SECOND) != 0 && cal.get(Calendar.MILLISECOND) != 0;
	}

	public FormatListener getFormatListener() {
		return formatListener;
	}

	public ExcelWriter setFormatListener(FormatListener formatListener) {
		this.formatListener = formatListener;
		return this;
	}

}
