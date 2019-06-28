package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JxlsComment {
	
	public static void execute() throws IOException {
		System.out.println("Inizio elaborazione...");
		InputStream in = new FileInputStream("src/main/resources/excel/comment_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/comment_output.xlsx");
		
		Workbook wb = new XSSFWorkbook(in);
		Sheet s = wb.getSheetAt(0);
		for (Row r : s)
			for (Cell c : r) {
				if (c.getCellType()==CellType.STRING && c.getStringCellValue().equals("comment here")) {
					System.out.println("Creo il commento...");
					CreationHelper factory = wb.getCreationHelper();
					ClientAnchor anchor = factory.createClientAnchor();
					anchor.setCol1(c.getColumnIndex()+1); // la box del commento parte qui...
					anchor.setCol2(c.getColumnIndex()+3); // ...e finisce qui
					anchor.setRow1(c.getRowIndex() + 1); // inizia una riga sotto la cella...
					anchor.setRow2(c.getRowIndex() + 5); // ...e 4 righe sopra
					
					@SuppressWarnings("rawtypes")
					Drawing drawing = s.createDrawingPatriarch();
					Comment comment = drawing.createCellComment(anchor);
					//testo e autore
					comment.setString(factory.createRichTextString("Hello"));
					comment.setAuthor("Me");
					
					c.setCellComment(comment);
					System.out.println("Commento applicato.");
				}
			}
		wb.write(out);
		wb.close();
		in.close();
		out.close();
		System.out.println("Fine elaborazione.");
	}
	
	public static void main(String[] args) throws IOException {
		execute();
	}

}
