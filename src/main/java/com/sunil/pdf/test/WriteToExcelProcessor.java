package com.sunil.pdf.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class WriteToExcelProcessor {
	
	@SuppressWarnings("unused")
	public void process(List<String> lines) throws IOException {

		String printText;
		
		// Create and write to excel sheet
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Testtab");
		XSSFFont font = ((XSSFWorkbook) workbook).createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 16);
		font.setBold(true);
		
		CellStyle style = workbook.createCellStyle();
		style.setWrapText(true);
		
		int rowCount = 2;
		
		String[][]  numberTextArray = new String[1000][1000];
		int rowNumber = 0, colNumber = 0;
		int previousColNumber = 0;
	
		for(String text : lines) {
			List<String> separatedText = Arrays.asList(text.split("\\|"));
			System.out.println(text);
			colNumber = previousColNumber;
			for(String s :  separatedText) {
				List<String> innerTextSeparatedText = Arrays.asList(s.split(" "));
				// Reset row count to zero so that 2D array doesn't have null values
				rowNumber = 0;				
				for(String numberText : innerTextSeparatedText) {
					printText = "";
					if(numberText.contains("-")) {
						printText = numberText.trim();
					} else {
						printText = printText.concat(" ").concat(numberText.trim());
					}
					
					System.out.println(printText);
					numberTextArray[rowNumber][colNumber] = printText;
					rowNumber++;
				}
				System.out.println("****");
				colNumber++;
			}
			// Save column number so that the row from file gets printed from the next column
			previousColNumber = colNumber;			
		}
		
		int noOfHeaderRows = 2, sheetRowCount = noOfHeaderRows;
		
		// Outer loop to limit the number of problems per row to 10
		for(int j = 0; j < colNumber; j=j+10) {
			sheetRowCount++;
			for(int i = 0; i < rowNumber; i++) {
				sheetRowCount = sheetRowCount + 1;
				Row row = sheet.createRow(sheetRowCount);
				int cellCount = 1;
				for(int innerJ = j; innerJ < j+10; innerJ++) {
					System.out.print(numberTextArray[i][innerJ] + " ");
					Cell cell = row.createCell(cellCount);
					cell.setCellValue(numberTextArray[i][innerJ]);
					cell.setCellStyle(style);
					cellCount++;
				}
				System.out.println();
			}
		}
		

		FileOutputStream outputStream = new FileOutputStream(new File("mantalmath.xlsx"));
		workbook.write(outputStream);
		workbook.close();
		
	}

	
}
