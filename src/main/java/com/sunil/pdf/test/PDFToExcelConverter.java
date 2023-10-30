package com.sunil.pdf.test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.pdfbox.pdmodel.PDDocument;

import technology.tabula.ObjectExtractor;
import technology.tabula.Page;
import technology.tabula.PageIterator;
import technology.tabula.RectangularTextContainer;
import technology.tabula.Table;
import technology.tabula.extractors.SpreadsheetExtractionAlgorithm;

public class PDFToExcelConverter {

	
	
	@SuppressWarnings({ "rawtypes", "resource" })
	public static void main(String[] args) throws IOException {
		WriteToExcelProcessor excelProcessor = new WriteToExcelProcessor();
		// InputStream in =
		// MainClass.class.getClassLoader().getResourceAsStream("abacus.pdf");
		File file = null;
		try {
			file = new File("mentalmath.pdf");
		} catch (Exception e) {
			System.out.println("error reading file");
			System.exit(1);
		}
		List<String> pdfRowsFromTable = new ArrayList<>();
		try (PDDocument document = PDDocument.load(file)) {
			SpreadsheetExtractionAlgorithm sea = new SpreadsheetExtractionAlgorithm();
			PageIterator pi = new ObjectExtractor(document).extract();
			while (pi.hasNext()) {
				// iterate over the pages of the document
				Page page = pi.next();
				List<Table> table = sea.extract(page);
				// iterate over the tables of the page
				for (Table tables : table) {
					List<List<RectangularTextContainer>> rows = tables.getRows();
					// iterate over the rows of the table
					for (List<RectangularTextContainer> cells : rows) {
						String printText = "";
						// print all column-cells of the row plus linefeed
						for (RectangularTextContainer content : cells) {
							// Note: Cell.getText() uses \r to concat text chunks
							String text = content.getText().replace("\r", " ");
							if (!text.isEmpty())
								printText = printText.concat(text + "|");
						}
						
						// Only print the rows you want in the excel
						if (!printText.isEmpty() && !(printText.contains("No.") || printText.contains("ANS")
								|| printText.contains("Ans") || printText.contains("Dist"))) {
							System.out.println(printText.substring(printText.indexOf("|")+1));
							pdfRowsFromTable.add(printText.substring(printText.indexOf("|")+1));
						}
					}
				}
			}
		}
		excelProcessor.process(pdfRowsFromTable);
	}

}
