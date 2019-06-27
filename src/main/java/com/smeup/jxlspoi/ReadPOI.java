package com.smeup.jxlspoi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadPOI {

	public static void main(String[] args) throws IOException {
		InputStream in = new FileInputStream("src/main/resources/excel/jxlspoi_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/jxlspoi_output.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> iRow = sheet.iterator();
		while (iRow.hasNext()) {
			Row nextRow = iRow.next();
			Iterator<org.apache.poi.ss.usermodel.Cell> iCell = nextRow.cellIterator();
			while (iCell.hasNext()) {
				Cell cell = iCell.next();
				System.out.print(cell.getAddress() + ": ");
				switch (cell.getCellType()) {
				case NUMERIC:
					System.out.println(cell.getNumericCellValue());
					break;
				case STRING:
					System.out.println(cell.getStringCellValue());
					break;
				case FORMULA:
					System.out.println(cell.getCellFormula());
					break;
				case BOOLEAN:
					System.out.println(cell.getBooleanCellValue());
					break;
				default:
					break;
				}
				System.out.println("Comment: "+cell.getCellComment());
			}
		}
		workbook.write(out);
		workbook.close();
		in.close();
		out.close();
		out.flush();
		System.out.println("Fine.");
	}

}
