package com.smeup.jxlspoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;

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
 * Prende un file Xml e lo esporta su Excel automaticamente,
 * senza che l'utente debba creare alcun file template
 */
public class ExportInExcel {

	@SuppressWarnings("rawtypes")
	public static void main(String[] args) throws IOException {
		File dir = new File("src/main/resources/xml/xmltest");
		boolean exists = false;
		// Prende il primo file con estensione .xml
		File[] list = dir.listFiles();
		File file = dir.listFiles()[0]; // Momentaneamente viene inizializzato come il primo elemento di dir
		for (File def : list)
			if (def.getName().contains(".xml")) {
				file = def;
				exists = true;
				break;
			}
		if (exists == false) {
			System.out.println("Nessun file con estensione .xml trovato");
			return;
		}
		OutputStream out = new FileOutputStream("src/main/resources/excel/export/export_input.xlsx");
		UIGridXmlObject u = new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(file));

		Workbook wb = new XSSFWorkbook();
		CreationHelper factory = wb.getCreationHelper();
		Sheet sheet = wb.createSheet("Template");
		Drawing drawing = sheet.createDrawingPatriarch();
		sheet.createRow(0);
		sheet.createRow(2);
		Row r = sheet.getRow(0);
		Row r1 = sheet.getRow(2);
		r.createCell(0);
		Cell c = r.getCell(0);
		c.setCellValue("Area");
		// Alla prima cella del primo foglio viene impostato il valore "Area".
		// Serve questa fase?

		wb.write(out);
		out.close();
		// Fine fase input, viene creato il file "export_input.xlsx"

		InputStream in = new FileInputStream("src/main/resources/excel/export/export_input.xlsx");
		OutputStream os = new FileOutputStream("src/main/resources/excel/export/export_temp.xlsx");
		Context context = new Context();
		int col = 1; // Numero della colonna
		for (UIGridColumn uc : u.getColumns()) {
			System.out.println("Aggiungo la colonna " + col);
			context.putVar("u1_col" + col, Arrays.asList(u.getFormattedColumnValues(uc.getCod())));
			col++;
		}
		CellAddress last = new CellAddress(0,0); // Modificare?

		for (int i = 0; i < u.getColumnsCount(); i++) {
			r1.createCell(i);
			Cell m = r1.getCell(i);
			System.out.println("M è all'indirizzo " + m.getAddress());
			m.setCellValue("${u1_col" + (i + 1) + "}");

			ClientAnchor anchor = factory.createClientAnchor();
			Comment comment = drawing.createCellComment(anchor);
			comment.setString(factory.createRichTextString("jx:each(lastCell='" + m.getAddress() + "' items='u1_col"
					+ (i + 1) + "' var='u1_col" + (i + 1) + "' direction='DOWN')"));
			comment.setAuthor("Me");

			m.setCellComment(comment);
			System.out.println("Commento applicato.");
			// Viene salvato l'indirizzo dell'ultima cella avente contenuto Jxls
			// così da potere definire la Xls Area
			if (last.getRow() < m.getRowIndex()) {
				int temp = last.getColumn();
				// Purtroppo non si possono settare singolarmente Row e Column
				last = new CellAddress(m.getRowIndex(), temp);
			}
			if (last.getColumn() < m.getColumnIndex()) {
				int temp = last.getRow();
				last = new CellAddress(temp, m.getColumnIndex());
			}
		}
		// Definizione area
		Cell ori = sheet.getRow(0).getCell(0);
		ClientAnchor anchor = factory.createClientAnchor();
		Comment comment = drawing.createCellComment(anchor);
		comment.setString(factory.createRichTextString("jx:area(lastCell='" + last + "')"));
		ori.setCellComment(comment);

		wb.write(os);
		wb.close();
		in.close();
		os.close();
		// Fine seconda fase, viene creato "export_temp.xlsx" contenente i comandi Jxls
		// con le note

		// Elaborazione JXLS
		System.out.println("Inizio elaborazione jxls...");
		InputStream inTemp = new FileInputStream("src/main/resources/excel/export/export_temp.xlsx");
		OutputStream outF = new FileOutputStream("src/main/resources/excel/export/export_output.xlsx");
		JxlsHelper.getInstance().processTemplate(inTemp, outF, context);
		inTemp.close();
		outF.close();
		System.out.println("Fine elaborazione.");
		// Fine ultima fase, viene creato "export_output.xlsx", che è il file finale
	}

}
