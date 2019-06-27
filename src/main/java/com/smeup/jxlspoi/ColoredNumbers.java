package com.smeup.jxlspoi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ColoredNumbers {

	public static void main(String[] args) throws IOException {
		System.out.println("Inizio elaborazione...");
		InputStream in = new FileInputStream("src/main/resources/excel/cpcpoi_output.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/cpcpoinumber_output.xlsx");
		Workbook workbook = new XSSFWorkbook(in);
		//Se dichiaro lo stile qua fuori me lo applica ai campi giusti
		/*CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setBold(true);
		style.setFont(font);*/
		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> iRow = sheet.rowIterator();
		ArrayList<Cell> listC = new ArrayList<>();
		while (iRow.hasNext()) {
			Row nextRow = iRow.next();
			Iterator<Cell> iCell = nextRow.cellIterator();
			while (iCell.hasNext()) {
				Cell nextCell = iCell.next();
				if (nextCell.getCellType()==CellType.NUMERIC)
					listC.add(nextCell);
			}
		}
		
		CellStyle numStyle = workbook.createCellStyle();
		Font numFont = workbook.createFont();
		numFont.setBold(true);
		numStyle.setFont(numFont);
		numStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		numStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		for (Cell i : listC)
			i.setCellStyle(numStyle);
		
		workbook.write(out);
		workbook.close();
		in.close();
		out.close();
		out.flush();
		System.out.println("Fine.");
	}

}
