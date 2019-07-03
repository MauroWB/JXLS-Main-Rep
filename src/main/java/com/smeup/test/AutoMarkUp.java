package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class AutoMarkUp {
	
	public static void read() throws IOException {
		InputStream in = new FileInputStream("src/main/resources/excel/auto/automu_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/auto/automu_temp.xlsx");
		Workbook wb = WorkbookFactory.create(in);
		for (Sheet s : wb)
			for (Row r : s)
				for (Cell c : r) {
					if (c.getCellType()==CellType.NUMERIC)
				}
		
		wb.write(out);
		out.close();
		wb.close();
		in.close();
	}
	
	public static void main(String[] args) {
		read();
		execute();
	}

}
