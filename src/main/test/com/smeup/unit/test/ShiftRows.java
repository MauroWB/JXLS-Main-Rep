package com.smeup.unit.test;

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class ShiftRows {
	
	@Test
	public void test1() throws EncryptedDocumentException, IOException {
		final InputStream in = new FileInputStream("src/main/resources/test/shift_template.xlsx");
		assertNotEquals(in, null);
		Workbook wb = WorkbookFactory.create(in);
		in.close();
		final OutputStream out = new FileOutputStream("src/main/resources/test/shift_output.xlsx");
		
		Sheet s = wb.getSheetAt(0);
		
		System.out.println(s.getRow(0).getCell(0).getStringCellValue());
		assertEquals(s.getRow(0).getCell(0).getStringCellValue(), "Shift Demo");
		assertNotEquals(wb.getSheetAt(0), null);
		assertNull("La quarta riga doveva essere null ma non lo è", s.getRow(3));
		for (int i = 0; i <= 5; i++) {
			if (s.getRow(i)==null) {
				s.createRow(i);
				System.out.println("Creo la riga numero "+i);
			}
			if (s.getRow(i).getCell(0)==null)
				s.getRow(i).createCell(0);
		}
		System.out.println(s.getRow(1).getCell(0).getAddress());
		s.shiftRows(0, 0, 1);
		assertEquals(s.getRow(1).getCell(0).getStringCellValue(), "Shift Demo");
		System.out.println(s.getRow(2).getCell(0).getAddress());
		System.out.println(s.getRow(1).getCell(0).getStringCellValue());
		assertEquals(s.getClass(), XSSFSheet.class);
		assertEquals(wb.getClass(), XSSFWorkbook.class);
		assertEquals(s.getRow(4).getClass(), XSSFRow.class);
		assertNotNull(s.getRow(5));
		
		wb.write(out);
		out.close();
		wb.close();
	}
	
}
