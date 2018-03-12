package com.anlohse.porter.test;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.junit.Assert;
import org.junit.Test;

import com.anlohse.porter.ExcelParser;
import com.anlohse.porter.ExcelParserException;
import com.anlohse.porter.ExcelWriter;
import com.anlohse.porter.FormatListener;

public class TestImport {

	@Test
	public void testImportXsl() throws IOException, ExcelParserException {
		List<Student> list = new ExcelParser().parseXls(TestImport.class.getResourceAsStream("/xsl/teste1.xls"),
				Student.class, 0, true);
		testa(list);
	}

	@Test
	public void testImportXslx() throws IOException, ExcelParserException {
		List<Student> list = new ExcelParser().parseXlsx(TestImport.class.getResourceAsStream("/xslx/teste1.xlsx"),
				Student.class, 0, true);
		testa(list);
	}

	@Test
	public void testImportArrayXsl() throws IOException, ExcelParserException {
		List<Object[]> list = new ExcelParser().parseXls(TestImport.class.getResourceAsStream("/xsl/teste1.xls"),
				0, true);
		testa2(list);
	}

	@Test
	public void testImportArrayXslx() throws IOException, ExcelParserException {
		List<Object[]> list = new ExcelParser().parseXlsx(TestImport.class.getResourceAsStream("/xslx/teste1.xlsx"),
				0, true);
		testa2(list);
	}

	@SuppressWarnings("deprecation")
	private void testa(List<Student> list) {
		Student st0 = list.get(0);
		Student st1 = list.get(1);
		Assert.assertEquals(st0.getName(), "Olswadino da Silva");
		Assert.assertEquals(st1.getName(), "Camila da Silva Correia");
		Assert.assertEquals(st0.getBirthDate(), new Date(107, 3, 23));
		Assert.assertEquals(st1.getBirthDate(), new Date(106, 10, 3));
		Assert.assertEquals(st0.getAge(), 9);
		Assert.assertEquals(st1.getAge(), 10);
		Assert.assertEquals(st0.getAvg(), 5.6, 0.001);
		Assert.assertEquals(st1.getAvg(), 6.8, 0.001);
	}

	@SuppressWarnings("deprecation")
	private void testa2(List<Object[]> list) {
		Object[] st0 = list.get(0);
		Object[] st1 = list.get(1);
		Assert.assertEquals(st0[0], "Olswadino da Silva");
		Assert.assertEquals(st1[0], "Camila da Silva Correia");
		Assert.assertEquals(st0[1], new Date(107, 3, 23));
		Assert.assertEquals(st1[1], new Date(106, 10, 3));
		Assert.assertEquals(st0[2], 9d);
		Assert.assertEquals(st1[2], 10d);
		Assert.assertEquals((Double) st0[3], 5.6, 0.001);
		Assert.assertEquals((Double) st1[3], 6.8, 0.001);
	}

	@Test
	public void testExportXsl() throws IOException, ExcelParserException {
		List<Student> sts = getdata();
		new ExcelWriter().writeXls(new FileOutputStream("tests/teste.xls"), sts, Student.class, null, true);
	}

	@Test
	public void testExportXslx() throws IOException, ExcelParserException {
		List<Student> sts = getdata();
		new ExcelWriter().setFormatListener(new TestFormatListener())
			.writeXlsx(new FileOutputStream("tests/teste.xlsx"), sts,Student.class, null, true);
	}

	@Test
	public void testExportArrayXsl() throws IOException, ExcelParserException {
		new ExcelWriter().writeXls(new FileOutputStream("tests/teste2.xls"), getarraydata(), null);
	}

	@Test
	public void testExportArrayXslx() throws IOException, ExcelParserException {
		new ExcelWriter().setFormatListener(new TestFormatListener())
			.writeXlsx(new FileOutputStream("tests/teste2.xlsx"), getarraydata(), null);
	}

	private List<Student> getdata() {
		@SuppressWarnings("deprecation")
		List<Student> sts = Arrays.asList(
				new Student("Albert Maxwell", new Date(105, 2, 11), 11, 6.7),
				new Student("Marie Einstein", new Date(104, 9, 26), 11, 4.6),
				new Student("Louis Lagrange", new Date(100, 0, 25), 12, 7.2));
		return sts;
	}

	private List<Object[]> getarraydata() {
		@SuppressWarnings("deprecation")
		List<Object[]> sts = Arrays.asList(
				new Object[] {"Nome", "Data Nasc.", "Idade", "Média"},
				new Object[] {"Albert Maxwell", new Date(105, 2, 11), 11, 6.7},
				new Object[] {"Marie Einstein", new Date(104, 9, 26), 11, 4.6},
				new Object[] {"Louis Lagrange", new Date(100, 0, 25), 12, 7.2});
		return sts;
	}

	private static int[] WIDTHS = {14*256, 20*256, 14*256, 14*256};
	
	private static class TestFormatListener implements FormatListener {
		
		private CellStyle headerStyle;
		private CellStyle redStyle;

		@Override
		public void formatHeader(Cell cell, String columnName, int c, int r, String header) {
			cell.setCellStyle(headerStyle);
		}

		@Override
		public void formatCell(Cell cell, String columnName, int c, int r, Object value) {
			if ((columnName != null && columnName.equals("avg")) || (c == 3 && r > 0)) {
				double v = (Double) value;
				if (v < 5) {
					cell.setCellStyle(redStyle);
				}
			}
			if (columnName == null && r == 0) {
				cell.setCellStyle(headerStyle);
			}
		}

		@Override
		public void start(Workbook wb, Sheet sheet) {
			XSSFFont font = (XSSFFont) wb.createFont();
			font.setFontName("Arial");
			font.setItalic(true);
			font.setBold(true);
			headerStyle = (XSSFCellStyle) wb.createCellStyle();
			headerStyle.setFont(font);
			headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
			font = (XSSFFont) wb.createFont();
			font.setFontName("Arial");
			font.setItalic(true);
			font.setColor(new XSSFColor(Color.red));
			redStyle = wb.createCellStyle();
		    redStyle.setFont(font);
			for (int i = 0; i < WIDTHS.length; i++) {
				sheet.setColumnWidth(i, WIDTHS[i]);
			}
		}

		@Override
		public void finish(Workbook wb, Sheet sheet) {
			
		}

		@Override
		public void formatRow(Row row) {
			// TODO Auto-generated method stub
			
		}
	};

}
