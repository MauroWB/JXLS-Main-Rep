package com.smeup.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
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

import Smeup.smeui.uidatastructure.uigridxml.UIGridColumn;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;


/*
 * Prende un file Xml e lo exporta su Excel automaticamente,
 * senza che venga creato alcun file template
 */
public class ExportInExcel {

	public static void main(String[] args) throws IOException {
		File dir = new File("src/main/resources/xml/export/");
		OutputStream out = new FileOutputStream("src/main/resources/excel/export/export_input.xlsx");

		List<UIGridXmlObject> master = new ArrayList<>();
		for (File f : dir.listFiles()) {
			master.add(new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(f)));
		}

		Workbook wb = new XSSFWorkbook();
		CreationHelper factory = wb.getCreationHelper();
		Sheet sheet = wb.createSheet("Template");
		sheet.createRow(0);
		sheet.createRow(2);
		Row r = sheet.getRow(0);
		Row r1 = sheet.getRow(2);
		r.createCell(0);
		Cell c = r.getCell(0);
		c.setCellValue("Area");

		wb.write(out);
		out.close();

		InputStream in = new FileInputStream("src/main/resources/excel/export/export_input.xlsx");
		OutputStream os = new FileOutputStream("src/main/resources/excel/export/export_temp.xlsx");
		Context context = new Context();
		int cont = 1;
		int col = 1;
		for (UIGridXmlObject u : master) {
			for (UIGridColumn lc : u.getColumns()) {
				System.out.println("Aggiungo la colonna " + col);
				context.putVar("u" + cont + "_col" + col, Arrays.asList(u.getFormattedColumnValues(lc.getCod())));
				col++;
			}
			cont++;
		}
		CellAddress last = new CellAddress(c);

		for (int i = 0; i < master.get(0).getColumnsCount(); i++) {
			r1.createCell(i);
			Cell m = r1.getCell(i);
			System.out.println("M è all'indirizzo "+m.getAddress());
			m.setCellValue("${u1_col" + (i + 1) + "}");

			ClientAnchor anchor = factory.createClientAnchor();
			anchor.setCol1(m.getColumnIndex() + 1); // la box del commento parte qui...
			anchor.setCol2(m.getColumnIndex() + 3); // ...e finisce qui
			anchor.setRow1(m.getRowIndex() + 1); // inizia una riga sotto la cella...
			anchor.setRow2(m.getRowIndex() + 5); // ...e 4 righe sopra

			@SuppressWarnings("rawtypes")
			Drawing drawing = sheet.createDrawingPatriarch();
			Comment comment = drawing.createCellComment(anchor);
			// testo e autore
			comment.setString(factory.createRichTextString("jx:each(lastCell='" + m.getAddress() + "' items='u1_col"
					+ (i + 1) + "' var='u1_col" + (i + 1) + "' direction='DOWN')"));
			comment.setAuthor("Me");

			m.setCellComment(comment);
			System.out.println("Commento applicato.");
			if (i == master.get(0).getColumnsCount()-1)
				last = m.getAddress();
		}
		//Definizione area
		Cell ori = sheet.getRow(0).getCell(0);
		ClientAnchor anchor = factory.createClientAnchor();
		anchor.setCol1(ori.getColumnIndex() + 1); // la box del commento parte qui...
		anchor.setCol2(ori.getColumnIndex() + 3); // ...e finisce qui
		anchor.setRow1(ori.getRowIndex() + 1); // inizia una riga sotto la cella...
		anchor.setRow2(ori.getRowIndex() + 5); // ...e 4 righe sopra
		@SuppressWarnings("rawtypes")
		Drawing drawing = sheet.createDrawingPatriarch();
		Comment comment = drawing.createCellComment(anchor);
		comment.setString(factory.createRichTextString("jx:area(lastCell='"+last+"')"));
		ori.setCellComment(comment);
		
		wb.write(os);
		wb.close();
		in.close();
		os.close();
		InputStream in1 = new FileInputStream("src/main/resources/excel/export/export_temp.xlsx");
		OutputStream out1 = new FileOutputStream("src/main/resources/excel/export/export_output.xlsx");
		JxlsHelper.getInstance().processTemplate(in1, out1, context);
		in1.close();
		out1.close();
	}

}
