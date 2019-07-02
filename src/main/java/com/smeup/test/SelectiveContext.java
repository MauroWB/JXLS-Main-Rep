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
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

/*
 * Legge i contenuti Excel e inserisce nel context
 * solo il necessario
 * 
 * WIP
 */
public class SelectiveContext {

	// Riempie la lista di tutti gli UIGridXmlObject relativi
	// a ciascun file .xml della directory specificata
	public static ArrayList<UIGridXmlObject> createData() {
		ArrayList<UIGridXmlObject> list = new ArrayList<>();
		int cont = 1;
		File dir = new File("src/main/resources/xml/xmltest");
		for (File f : dir.listFiles()) {
			if (f.getName().endsWith(".xml")) {
				UIGridXmlObject u = new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(f));
				u.setComment("u" + cont); // Commento inteso come nome della tabella.
				// Servirà più avanti per identificare la tabella
				System.out.println(u.getComment());
				list.add(u);
				cont++;
			}
		}
		return list;
	}

	/*
	 * Legge il contenuto delle celle, elabora di modo tale che abbia in mano un
	 * oggetto (se esiste) UIGridXmlObject, lo mette nel context (se non c'è già)
	 */
	public static Context readStuff(List<UIGridXmlObject> list) throws Exception, IOException {
		System.out.println("Sono dentro");
		InputStream in = new FileInputStream("src/main/resources/excel/sel_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/sel_temp.xlsx");
		Context context = new Context();
		Workbook wb = WorkbookFactory.create(in);
		CellAddress lastCell = new CellAddress(0, 0);
		CreationHelper factory = wb.getCreationHelper(); // Occorrerà per il commento
		for (Sheet s : wb) {
			@SuppressWarnings("rawtypes")
			Drawing drawing = s.createDrawingPatriarch(); // Occorrerà per il commento
			for (Row r : s) {
				for (Cell c : r) {
					if (c.getStringCellValue().startsWith("${") && c.getStringCellValue().endsWith("}")) {
						// Quindi è un comando Jxls
						// Toglie i caratteri speciali dalla cella
						String string = c.getStringCellValue().replace("${", "");
						string = string.replace("}", "");
						System.out.println(" Cella " + c.getAddress() + ": " + string);
						String arrString[] = string.split("_");

						// Cerca se esiste nella lista un uxo con commento = arrString[0]
						for (UIGridXmlObject uxo : list) {
							if (uxo.getComment().equals(arrString[0])) {
								boolean flag = true;
								// Inserisce l'uxo se non esiste nel context
								if (context.getVar(arrString[0] + "_" + arrString[1]) == null) {
									StringBuffer col = new StringBuffer();
									for (int i = 0; i < arrString[1].length(); i++) {
										if (Character.isDigit(arrString[1].charAt(i))) {
											col.append(arrString[1].charAt(i));
										}
									}
									int numCol = Integer.parseInt(col.toString());

									try {
										System.out.println(
												"  Inserisco nel context " + arrString[0] + "_" + arrString[1]);
										context.putVar(arrString[0] + "_" + arrString[1], Arrays.asList(
												uxo.getFormattedColumnValues(uxo.getColumnByIndex(numCol).getCod())));
										// Passata come Arrays.asList, sennò non si può iterare in Jxls
									} catch (Exception e) {
										System.out.println("  Non esiste la colonna indicata.");
										e.printStackTrace();
										flag = false;
										// Se la colonna non esiste, non c'è bisogno che venga applicato il commento in
										// Excel, quindi flag va a false
									}

								} else {
									System.out.println(
											"  Esista già nel context l'oggetto " + arrString[0] + "_" + arrString[1]);
								}

								if (flag) {
									// Crea e applica il commento
									// I valori di sotto sono da specificare, sennò tutti i commenti vengono
									// applicati nella stessa cella e va in errore
									ClientAnchor anchor = factory.createClientAnchor();
									anchor.setCol1(c.getColumnIndex() + 1);
									anchor.setCol2(c.getColumnIndex() + 3);
									anchor.setRow1(c.getRowIndex() + 1);
									anchor.setRow2(c.getRowIndex() + 5);
									Comment comment = drawing.createCellComment(anchor);
									comment.setString(factory.createRichTextString("jx:each(lastCell='" + c.getAddress()
											+ "' items='" + arrString[0] + "_" + arrString[1] 
											+ "' var='" + arrString[0] + "_" + arrString[1] 
											+ "' direction='DOWN'" + ")"));
									comment.setAuthor("Author");
									c.setCellComment(comment);
									lastCell = c.getAddress();
								}
							}
						}
					}
				}
			}
			// Crea area
			Cell origin = s.getRow(0).getCell(0); // Area che parte dalla prima cella del foglio
			ClientAnchor anchor = factory.createClientAnchor();
			anchor.setCol1(origin.getColumnIndex() + 1);
			anchor.setCol2(origin.getColumnIndex() + 3);
			anchor.setRow1(origin.getRowIndex() + 1);
			anchor.setRow2(origin.getRowIndex() + 5);
			Comment comment = drawing.createCellComment(anchor);
			comment.setString(factory.createRichTextString("jx:area(lastCell='"+lastCell+"')"));
			comment.setAuthor("Author");
			origin.setCellComment(comment);
		}
		wb.write(out);
		out.close();
		in.close();
		wb.close();
		return context;
	}

	public static void main(String[] args) throws Exception {
		System.out.println("Inizio elaborazione...");
		InputStream in = new FileInputStream("src/main/resources/excel/sel_temp.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/sel_output.xlsx");
		Context context = new Context();
		List<UIGridXmlObject> list = createData();
		context = readStuff(list);
		JxlsHelper.getInstance().processTemplate(in, out, context);
		in.close();
		out.close();
		System.out.println("Fine elaborazione.");
	}

}
