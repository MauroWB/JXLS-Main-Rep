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
		for (Row r : s)
			for (Cell c : r) {
				if (c.getCellType() == CellType.STRING) {
					if (c.getStringCellValue().equals("comment here")) {
						System.out.println("Creo il commento...");

						ClientAnchor anchor = factory.createClientAnchor();
						anchor.setCol1(c.getColumnIndex() + 1); // la box del commento parte qui...
						anchor.setCol2(c.getColumnIndex() + 3); // ...e finisce qui
						anchor.setRow1(c.getRowIndex() + 1); // inizia una riga sotto la cella...
						anchor.setRow2(c.getRowIndex() + 5); // ...e 4 righe sopra

						@SuppressWarnings("rawtypes")
						Drawing drawing = s.createDrawingPatriarch();
						Comment comment = drawing.createCellComment(anchor);
						// testo e autore
						comment.setString(factory.createRichTextString("Hello"));
						comment.setAuthor("Me");

						c.setCellComment(comment);
						System.out.println("Commento applicato.");
					}
					if (c.getStringCellValue().startsWith("PRINT")) {
						String stuff[] = c.getStringCellValue().split(" ");
						context.putVar(stuff[1], Arrays.asList(new String[] { "Test", "Test2" }));
						ClientAnchor anchor = factory.createClientAnchor();
						/*
						 * Occorre sempre definire queste coordinate, in quanto più commenti
						 * non possono essere nella stessa posizione
						 */
						anchor.setCol1(c.getColumnIndex() + 1); // la box del commento parte qui...
						anchor.setCol2(c.getColumnIndex() + 3); // ...e finisce qui
						anchor.setRow1(c.getRowIndex() + 1); // inizia una riga sotto la cella...
						anchor.setRow2(c.getRowIndex() + 5); // ...e 4 righe sopra

						@SuppressWarnings("rawtypes")
						Drawing drawing = s.createDrawingPatriarch();
						Comment comment = drawing.createCellComment(anchor);
						// testo e autore
						comment.setString(factory.createRichTextString("jx:each(lastCell=" + "'" + c.getAddress() + "'"
								+ " items=" + "'" + stuff[1] + "'" + " var=" + "'obj')"));
						comment.setAuthor("Me");
						c.setCellComment(comment);
						c.setCellValue("${obj}");
					}
					/*
					 * if (c.getStringCellValue().toLowerCase().equals("start")); { ClientAnchor
					 * anchor = factory.createClientAnchor(); anchor.setCol1(c.getColumnIndex() +
					 * 1); anchor.setCol2(c.getColumnIndex() + 3); anchor.setRow1(c.getRowIndex() +
					 * 1); anchor.setRow2(c.getRowIndex() + 5);
					 * 
					 * Drawing drawing = s.createDrawingPatriarch(); Comment comment =
					 * drawing.createCellComment(anchor);
					 * 
					 * //TODO Generalizzare di modo tale che trovi da solo l'area di fine
					 * comment.setString(factory.createRichTextString("jx:area(lastCell='A7')"));
					 * c.setCellComment(comment); }
					 */
				}
			}

		wb.write(out);
		wb.close();
		in.close();
		out.close();
		System.out.println("Fine elaborazione.");
		System.out.println("Inizio processing con Jxls...");

		/* Test */
		InputStream is = new FileInputStream("src/main/resources/excel/comment_output.xlsx");
		OutputStream os = new FileOutputStream("src/main/resources/excel/commentelab_output.xlsx");
		JxlsHelper.getInstance().processTemplate(is, os, context);
		is.close();
		os.close();
		System.out.println("Fine processing.");
		/* Fine Test */
	}

	public static void main(String[] args) throws IOException {
		execute();
		// process();
	}

}
