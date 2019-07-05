package com.smeup.jxlspoi;

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

import com.smeup.test.POIUtilities;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

/*
 * Legge i contenuti Excel e inserisce nel context
 * solo il necessario
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

	// Legge il contenuto delle celle, elabora di modo tale che abbia in mano un
	// oggetto (se esiste) UIGridXmlObject, lo mette nel context (se non c'è già)
	public static Context read(List<UIGridXmlObject> list) throws Exception, IOException {
		InputStream in = new FileInputStream("src/main/resources/excel/sel_cont/sel_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/sel_cont/sel_temp.xlsx");
		Context context = new Context();
		Workbook wb = WorkbookFactory.create(in);
		CellAddress lastCell = new CellAddress(0, 0);
		CreationHelper factory = wb.getCreationHelper(); // Occorrerà per creare i commenti
		for (Sheet s : wb) {
			@SuppressWarnings("rawtypes")
			Drawing drawing = s.createDrawingPatriarch(); // Occorrerà per creare i commenti
			for (Row r : s) {
				for (Cell c : r) {
					if (c.getStringCellValue().startsWith("${") && c.getStringCellValue().endsWith("}")) {
						// (Quindi se è un comando Jxls)
						// Toglie i caratteri speciali dalla cella
						String string = c.getStringCellValue().replace("${", "");
						string = string.replace("}", "");
						System.out.println(" Cella " + c.getAddress() + ": " + string);
						String arrString[] = string.split("_");

						// Cerca se esiste nella lista un uxo con commento == arrString[0]
						for (UIGridXmlObject uxo : list) {
							if (uxo.getComment().equals(arrString[0])) {

								boolean flag = true;
								// Inserisce l'uxo nel context solo se non è stato ancora inserito
								if (context.getVar(arrString[0] + "_" + arrString[1]) == null) {
									StringBuffer col = new StringBuffer();
									for (int i = 0; i < arrString[1].length(); i++) {
										if (Character.isDigit(arrString[1].charAt(i))) {
											col.append(arrString[1].charAt(i));
										}
									}
									int numCol = Integer.parseInt(col.toString());
									// Dentro numCol c'è, appunto, il numero della colonna

									try {
										System.out.println(
												"  Inserisco nel context " + arrString[0] + "_" + arrString[1]);
										// Volendo passare la lista delle colonne + intestazione al primo posto si usa
										// il metodo add (int, obj), che fa shiftare tutti gli elementi
										List<Object> arr = new ArrayList<>();
										arr.add(uxo.getColumnByIndex(numCol).getTxt());
										arr.addAll(Arrays.asList(
												uxo.getFormattedColumnValues(uxo.getColumnByIndex(numCol).getCod())));
										context.putVar(arrString[0] + "_" + arrString[1], arr);
										// Passata come Arrays.asList, sennò non si può iterare in Jxls
									} catch (Exception e) {
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
									ClientAnchor anchor = factory.createClientAnchor();
									Comment comment = drawing.createCellComment(anchor);
									comment.setString(factory.createRichTextString("jx:each(lastCell='" + c.getAddress()
											+ "' items='" + arrString[0] + "_" + arrString[1] + "' var='" + arrString[0]
											+ "_" + arrString[1] + "' direction='DOWN'" + ")"));
									comment.setAuthor("Author");
									c.setCellComment(comment);

									// Per dare le dimensioni giuste alla XLS Area
									if (lastCell.getRow() < c.getRowIndex()) {
										int temp = lastCell.getColumn();
										// Purtroppo non si possono settare singolarmente Row e Column
										lastCell = new CellAddress(c.getRowIndex(), temp);
									}
									if (lastCell.getColumn() < c.getColumnIndex()) {
										int temp = lastCell.getRow();
										lastCell = new CellAddress(temp, c.getColumnIndex());
									}
								}
							}
						}
					}
				}
			}
			// Ora che l'ultima cella interessata è stata definita, si definisce l'area
			ClientAnchor anchor = factory.createClientAnchor();
			Comment comment = drawing.createCellComment(anchor);
			comment.setString(factory.createRichTextString("jx:area(lastCell='" + lastCell + "')"));
			comment.setAuthor("Author");
			comment.setAddress(0, 0); // Viene applicato alla prima cella del foglio
		}
		wb.write(out);
		out.close();
		in.close();
		wb.close();
		return context;
	}

	public static void adjust() throws IOException {
		System.out.println(" Formattazione colonne...");
		InputStream in = new FileInputStream("src/main/resources/excel/sel_cont/sel_output.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/sel_cont/sel_formattedoutput.xlsx");
		Workbook wb = WorkbookFactory.create(in);
		POIUtilities.fitColumns(wb);
		wb.write(out);
		wb.close();
		in.close();
		out.close();
		System.out.println(" Fine formattazione colonne.");
	}

	public static void main(String[] args) throws Exception {
		System.out.println("Inizio elaborazione...");
		InputStream in = new FileInputStream("src/main/resources/excel/sel_cont/sel_temp.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/sel_cont/sel_output.xlsx");
		Context context = new Context();
		List<UIGridXmlObject> list = createData();
		context = read(list);
		JxlsHelper.getInstance().processTemplate(in, out, context);
		in.close();
		out.close();
		adjust();
		System.out.println("Fine elaborazione.");
	}

}
