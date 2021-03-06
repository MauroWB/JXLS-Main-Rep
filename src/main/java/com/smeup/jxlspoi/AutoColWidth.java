package com.smeup.jxlspoi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AutoColWidth {

	public static void main(String[] args) throws IOException {
		System.out.println("Inizio elaborazione...");
		InputStream in = new FileInputStream("src/main/resources/excel/cpcpoi_output.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/cpcpoiautotest_output.xlsx");
		Workbook workbook = new XSSFWorkbook(in);
		Sheet sheet = workbook.getSheetAt(0);
		for (Row r : sheet)
			for (Cell c : r)
				sheet.autoSizeColumn(c.getColumnIndex());
		workbook.write(out);
		workbook.close();
		in.close();
		out.close();
		System.out.println("Fine elaborazione.");
	}

}
