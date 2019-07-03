package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

/*
 * Crea e applica un commento a
 * una cella che presenta un testo
 * specifico.
 */

public class JxlsCommentDemo {

	public static void execute() throws IOException {
		System.out.println("Inizio elaborazione...");
		InputStream in = new FileInputStream("src/main/resources/excel/comment_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/comment_output.xlsx");
		Context context = new Context();

		Workbook wb = new XSSFWorkbook(in);
		CreationHelper factory = wb.getCreationHelper();
		Sheet s = wb.getSheetAt(0);
		@SuppressWarnings("rawtypes")
		Drawing drawing = s.createDrawingPatriarch();
		CellAddress last = s.getRow(0).getCell(0).getAddress();
		for (Row r : s) {
			for (Cell c : r) {
				if (c.getCellType() == CellType.STRING) {
					System.out.println("Sono alla cella "+c.getAddress());
					if (c.getStringCellValue().equals("comment here")) {
						System.out.println("Creo il commento...");
						ClientAnchor anchor = factory.createClientAnchor();
						Comment comment = drawing.createCellComment(anchor);
						// testo e autore
						comment.setString(factory.createRichTextString("Hello"));
						comment.setAuthor("Me");
						c.setCellComment(comment);
						System.out.println("Commento applicato.");
					}
					else if (c.getStringCellValue().startsWith("PRINT")) {
						String stuff[] = c.getStringCellValue().split(" ");
						context.putVar(stuff[1], Arrays.asList(new String[] { "Test", "Test2" }));
						ClientAnchor anchor = factory.createClientAnchor();
						Comment comment = drawing.createCellComment(anchor);
						// testo e autore
						comment.setString(factory.createRichTextString("jx:each(lastCell=" + "'" + c.getAddress() + "'"
								+ " items=" + "'" + stuff[1] + "'" + " var=" + "'obj')"));
						comment.setAuthor("Me");
						c.setCellComment(comment);
						c.setCellValue("${obj}");
						last = c.getAddress();
					}
				}
			}
		}
		ClientAnchor anchor = factory.createClientAnchor();
		Comment comment = drawing.createCellComment(anchor);
		comment.setString(factory.createRichTextString("jx:area(lastCell='"+last+"')"));
		comment.setAddress(0, 0);

		wb.write(out);
		wb.close();
		in.close();
		out.close();
		System.out.println("Fine elaborazione.");
		System.out.println("Inizio processing con Jxls...");
		InputStream is = new FileInputStream("src/main/resources/excel/comment_output.xlsx");
		OutputStream os = new FileOutputStream("src/main/resources/excel/commentelab_output.xlsx");
		JxlsHelper.getInstance().processTemplate(is, os, context);
		is.close();
		os.close();
		System.out.println("Fine processing.");
	}

	public static void main(String[] args) throws IOException {
		execute();
		// process();
	}

}
