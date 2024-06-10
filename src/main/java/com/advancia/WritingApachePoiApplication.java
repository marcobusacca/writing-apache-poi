package com.advancia;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingApachePoiApplication {
	
	private static final String FILE_PATH = "C:/Users/marco/_Advancia/Formazione/1-JAVA/APACHE-POI-AND-ITEXT/writing-apache-poi/file.xlsx";

	public static void main(String[] args) {

		try (Workbook workbook = new XSSFWorkbook()) {

			Object[] headingSet = { "Codice", "Nome", "Quantità in magazzino", "Quantità venduta", "Quantità totale" };

			Object[][] dataSet = { { "001", "Latte", 1, 2, 0 }, { "002", "Miele", 4, 3, 0 },
					{ "003", "Pane", 5, 10, 0 }, { "004", "Pasta", 6, 4, 0 }, { "005", "Carne", 2, 3, 0 } };
			
			writeNewSheet(workbook, "Lista prodotti", headingSet, dataSet);
			writeFile(workbook);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void writeNewSheet(Workbook workbook, String sheetName, Object[] headingSet, Object[][] dataSet) {
		
		Sheet sheet = workbook.createSheet(sheetName);

		writeHeadingSetInSheet(workbook, sheet, headingSet);
		writeDataSetInSheet(workbook, sheet, dataSet);
	}

	public static void writeHeadingSetInSheet(Workbook workbook, Sheet sheet, Object[] headingSet) {

		Row row = sheet.createRow(0);

		int cellIndex = 0;

		for (Object headingCell : headingSet) {

			Cell cell = row.createCell(cellIndex++);

			cell.setCellValue((String) headingCell);

			cell.setCellStyle(getCellStyle(workbook, true));

			sheet.autoSizeColumn(cellIndex - 1);
		}
	}

	public static void writeDataSetInSheet(Workbook workbook, Sheet sheet, Object[][] dataSet) {

		int rowIndex = 1;

		for (Object[] dataRow : dataSet) {

			Row row = sheet.createRow(rowIndex++);

			int cellIndex = 0;

			for (Object dataCell : dataRow) {

				Cell cell = row.createCell(cellIndex++);

				switch (cellIndex) {
					case 1:
					case 2:
						cell.setCellValue((String) dataCell);
						break;
					case 3:
					case 4:
						cell.setCellValue((Integer) dataCell);
						break;
					case 5:
						cell.setCellFormula("SUM(C" + rowIndex + ":D" + rowIndex + ")");
						break;
					default:
						break;
				}

				cell.setCellStyle(getCellStyle(workbook, false));

				sheet.autoSizeColumn(cellIndex - 1);
			}
		}
	}

	public static CellStyle getCellStyle(Workbook workbook, boolean isWithForegroundColor) {
		
		CellStyle style = workbook.createCellStyle();
		
		if (isWithForegroundColor) {
			style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}
		
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		
		style.setBorderTop(BorderStyle.THICK);
		style.setBorderBottom(BorderStyle.THICK);
		style.setBorderLeft(BorderStyle.THICK);
		style.setBorderRight(BorderStyle.THICK);
		
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());

		Font font = workbook.createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 30);
		style.setFont(font);

		return style;
	}

	public static void writeFile(Workbook workbook) {
		try (FileOutputStream outputStream = new FileOutputStream(FILE_PATH)) {
			workbook.write(outputStream);
		} catch (Exception e) { // FileNotFoundException || IOException
			e.printStackTrace();
		}
	}
}
