package com.smeup.jxlspoi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellUtil;

/*
 * Usa dei loop for each anziché gli iteratori di ColoredNumbers
 * 
 * Si usa setCellProperties in quanto consente di
 * non dovere creare nuovi stili, tuttavia ci sono stati 
 * malfunzionamenti con il font
 */

public class ThinColoredNumbers {

	public static void main(String[] args) throws IOException {
		InputStream in = new FileInputStream("src/main/resources/excel/cpcpoi_output.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/cpcpoinumber_output.xlsx");
		Workbook wb = WorkbookFactory.create(in);
		for (Sheet sheet : wb)
			for (Row row : sheet)
				for (Cell cell : row)
					switch (cell.getCellType()) {
						case NUMERIC:
							CellUtil.setCellStyleProperty(cell, CellUtil.FILL_FOREGROUND_COLOR, IndexedColors.GREY_50_PERCENT.getIndex());
						default:
							break;
					}
		wb.write(out);
		wb.close();
		in.close();
		out.close();
		out.flush();
		System.out.println("Fine.");
	}

}
