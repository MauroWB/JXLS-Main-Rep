package com.smeup.jxlspoi;

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
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

/*
 * Crea e applica un commento a
 * una cella che presenta un testo
 * specifico.
 */

public class JxlsCommentDemo {

	@SuppressWarnings("rawtypes")
	public static void execute() throws IOException {
		System.out.println("Inizio elaborazione...");
		InputStream in = new FileInputStream("src/main/resources/excel/comment_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/comment_output.xlsx");
		Context context = new Context();
		Workbook wb = new XSSFWorkbook(in);
		Sheet s = wb.getSheetAt(0);
		
		CreationHelper factory = wb.getCreationHelper();
		Drawing drawing = s.createDrawingPatriarch();
		
		for (Row r : s) {
			for (Cell c : r) {
				if (c.getCellType() == CellType.STRING && c.getStringCellValue().equals("comment here")) {
					System.out.println("Creo il commento...");
					ClientAnchor anchor = factory.createClientAnchor();
					Comment comment = drawing.createCellComment(anchor);
					// Testo e autore
					comment.setString(factory.createRichTextString("Hello"));
					comment.setAuthor("Me");
					c.setCellComment(comment);
					System.out.println("Commento applicato.");
				}
			}
		}

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
