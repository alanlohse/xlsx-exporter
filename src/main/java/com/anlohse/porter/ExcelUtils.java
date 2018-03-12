package com.anlohse.porter;

import java.awt.Color;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

	public static void insertImage(Workbook workbook, Sheet sheet, int row1, int col1, int row2, int col2, byte[] imageBytes, int imageType) {
		   int pictureureIdx = workbook.addPicture(imageBytes, imageType);
		   CreationHelper helper = workbook.getCreationHelper();
		   Drawing drawing = sheet.createDrawingPatriarch();
		   ClientAnchor anchor = helper.createClientAnchor();
		   anchor.setCol1(col1);
		   anchor.setRow1(row1);
		   anchor.setCol2(col2);
		   anchor.setRow2(row2);
		   drawing.createPicture(anchor, pictureureIdx);
	}
	
	public static void mergeCells(Sheet sheet, int row1, int col1, int row2, int col2) {
		sheet.addMergedRegion(new CellRangeAddress(row1, row2, col1, col2));		
	}
	
	public static CellStyle createStyle(Workbook workbook, String fontName, double fSize, boolean bold, boolean italic, HorizontalAlignment hAlign, VerticalAlignment vAlign) {
		CellStyle style = workbook.createCellStyle();
		if (hAlign != null)
			style.setAlignment(hAlign);
		if (vAlign != null)
			style.setVerticalAlignment(vAlign);
		if (fontName != null) {
			Font font = workbook.createFont();
			font.setFontName(fontName);
			font.setFontHeight((short)(fSize * 20));
			font.setBold(bold);
			font.setItalic(italic);
			style.setFont(font);
		}
		return style;
	}

	public static void setFillBackgroundColor(Workbook wb, CellStyle style, int r, int g, int b) {
		if (wb instanceof XSSFWorkbook)
			setXSSFFillBackgroundColor((XSSFWorkbook) wb, (XSSFCellStyle) style, r, g, b);
		else
			setHSSFFillBackgroundColor((HSSFWorkbook) wb, (HSSFCellStyle) style, r, g, b);
	}

	private static void setXSSFFillBackgroundColor(XSSFWorkbook wb, XSSFCellStyle style, int r, int g, int b) {
		style.setFillBackgroundColor(new XSSFColor(new Color(r,g,b)));
	}
	private static void setHSSFFillBackgroundColor(HSSFWorkbook wb, HSSFCellStyle style, int r, int g, int b) {
		HSSFPalette palette = wb.getCustomPalette();
		HSSFColor color = palette.findColor((byte) r, (byte) g, (byte) b);
		if (color == null) {
			color = palette.addColor((byte) r, (byte) g, (byte) b);
		}
		style.setFillBackgroundColor(color.getIndex());
	}

	public static void setFillForegroundColor(Workbook wb, CellStyle style, int r, int g, int b) {
		if (wb instanceof XSSFWorkbook)
			setXSSFFillForegroundColor((XSSFWorkbook) wb, (XSSFCellStyle) style, r, g, b);
		else
			setHSSFFillForegroundColor((HSSFWorkbook) wb, (HSSFCellStyle) style, r, g, b);
	}

	private static void setXSSFFillForegroundColor(XSSFWorkbook wb, XSSFCellStyle style, int r, int g, int b) {
		style.setFillForegroundColor(new XSSFColor(new Color(r,g,b)));
	}
	private static void setHSSFFillForegroundColor(HSSFWorkbook wb, HSSFCellStyle style, int r, int g, int b) {
		HSSFPalette palette = wb.getCustomPalette();
		HSSFColor color = palette.findColor((byte) r, (byte) g, (byte) b);
		if (color == null) {
			color = palette.addColor((byte) r, (byte) g, (byte) b);
		}
		style.setFillForegroundColor(color.getIndex());
	}

	public static final short pt2(double pts) {
		return (short) (pts * 20.0);		
	}
	
}
