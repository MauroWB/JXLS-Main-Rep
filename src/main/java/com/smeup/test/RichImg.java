package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
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

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class RichImg {

	/*public static void readMatrix(UIGridXmlObject uxo) throws IOException {
		System.out.println("Inizio.");
		InputStream in = new FileInputStream("src/main/resources/excel/img/richimg_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/img/richimg_tempo.xlsx");

		System.out.println("Uxo creato.");
		int index = 0; // In assenza del metodo getColumnIndex
		for (UIGridColumn uc : uxo.getColumns()) {
			System.out.println("Scorro le colonne.");
			if (uc.getOgg().equals("J4IMG")) {
				System.out.println("Sono nel caso.");
				System.out.println(uc.getTipoSmeup());
				String url = SmeupImageUtility.getImageUrl(uc.getTipoSmeup(), uc.getParametroSmeup(), uc.getCod(), null);
				//SmeupImageUtility.getImageUrl(aTipo, aParametro, aCodice, aIndex);
				//System.out.println("Url:"+url);
				// salva numero colonna, così che l'elemento index di ogni riga sia flaggato
				// come immagine

				index++;
			}
		}

		Context context = new Context();
		ArrayList<ArrayList<Object>> master = new ArrayList<ArrayList<Object>>();
		for (int i = 0; i < uxo.getColumnsCount(); i++) {
			ArrayList<Object> sub = new ArrayList<>();
			sub.addAll(Arrays.asList(uxo.getFormattedColumnValues(uxo.getColumnByIndex(i).getCod())));
			master.add(sub);
		}
		context.putVar("master", master);
		JxlsHelper.getInstance().processTemplate(in, out, context);
		in.close();
		out.close();
	}*/

	public static void applyComments(UIGridXmlObject uxo) throws IOException {
		// Dopo prima elaborazione Jxls?
		OutputStream out = new FileOutputStream("src/main/resources/excel/img/richimg_tempo.xlsx");
		Workbook wb = new XSSFWorkbook();
		Sheet s = wb.createSheet("Sheet1");
		CreationHelper factory = wb.getCreationHelper();
		Drawing<?> drawing = s.createDrawingPatriarch();
		@SuppressWarnings("unused")
		Row r1 = s.createRow(0); // Va creata affinché ci si possa riferire alla cella (0,0)
		Row r = s.createRow(2);
		CellAddress lastCell = new CellAddress(0, 0);

		for (int i = 0; i < uxo.getColumnsCount(); i++) {
			Cell c = r.createCell(i);
			c.setCellValue("${col}");
			ClientAnchor anchor = factory.createClientAnchor();
			Comment comment = drawing.createCellComment(anchor);
			comment.setString(factory.createRichTextString(
					"jx:each(direction='DOWN' lastCell='" + new CellAddress(c.getRowIndex() + 1, c.getColumnIndex())
							+ "' items='col" + (i + 1) + "' var='col')"));
			c.setCellComment(comment);
			lastCell = c.getAddress(); // In questo esempio funziona anche così
		}
		ClientAnchor anchor = factory.createClientAnchor();
		Comment comment = drawing.createCellComment(anchor);
		comment.setString(factory.createRichTextString(
				"jx:area(lastCell='" + new CellAddress(lastCell.getRow() + 1, lastCell.getColumn()) + "')"));
		comment.setAddress(0, 0); // può dare errore per cella non creata

		wb.write(out);
		out.close();
		wb.close();
	}

	public static void process(UIGridXmlObject uxo) throws IOException {
		FileInputStream inTemp = new FileInputStream("src/main/resources/excel/img/richimg_tempo.xlsx");
		FileOutputStream outF = new FileOutputStream("src/main/resources/excel/img/richimg_output.xlsx");
		Context context = new Context();
		for (int i = 0; i < uxo.getColumnsCount(); i++) {
			ArrayList<Object> sub = new ArrayList<>();
			sub.addAll(Arrays.asList(uxo.getFormattedColumnValues(uxo.getColumnByIndex(i).getCod())));
			sub.addAll(Arrays.asList(uxo.getColumnValues(i)));
			context.putVar("col" + (i + 1), sub);
		}
		System.out.println(context.toString());
		JxlsHelper.getInstance().processTemplate(inTemp, outF, context);
		inTemp.close();
		outF.close();
	}

	public static void main(String[] args) throws IOException {
		UIGridXmlObject uxo = new UIGridXmlObject(
				UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/xml/richimg.xml"));
		// Leggi matrice, scorre colonne: if colonna.ogg=J4IMG, processala con SmeIMG
	//readMatrix(uxo);
		applyComments(uxo);
		process(uxo);
	}

}
