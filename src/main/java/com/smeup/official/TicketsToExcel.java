package com.smeup.official;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellUtil;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.jxls.util.Util;

import Smeup.smeui.uidatastructure.uigridxml.UIGridRow;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class TicketsToExcel {

	/**
	 * Carica le immagini nel context come ByteArray.
	 * 
	 * @param context - Il context in cui vanno inserite le immagini.
	 * @return context con immagini inserite.
	 * @throws IOException
	 */
	public static Context loadImages(Context context) throws IOException {
		InputStream imgInputStream = new FileInputStream("src/main/resources/ticket/img/static_image.png");
		byte[] imageBytes = Util.toByteArray(imgInputStream);
		InputStream userIS = new FileInputStream("src/main/resources/ticket/img/user_placeholder.png");
		byte[] userBytes = Util.toByteArray(userIS);

		context.putVar("static_image", imageBytes);
		context.putVar("user_ph", userBytes);

		userIS.close();
		imgInputStream.close();
		return context;
	}

	/**
	 * In base al contenuto della cella "Urgenza", cambia il colore della barra di
	 * ciascuna entry.
	 * 
	 * @param wb - Workbook su cui agire.
	 * @return Workbook con colori relativi a ciascun campo urgenza.
	 */
	public static Workbook urgencyColor(Workbook wb, int startCol) {

		Font bassa = wb.createFont();
		bassa.setColor(IndexedColors.SEA_GREEN.getIndex());
		bassa.setBold(true);
		bassa.setFontName("Lato UI");
		
		Font media = wb.createFont();
		media.setColor(IndexedColors.LIGHT_ORANGE.getIndex());
		media.setBold(true);
		media.setFontName("Lato UI");
		
		Font alta = wb.createFont();
		alta.setColor(IndexedColors.RED.getIndex());
		alta.setBold(true);
		alta.setFontName("Lato UI");
		
		Font urgente = wb.createFont();
		urgente.setColor(IndexedColors.BLACK.getIndex());
		urgente.setBold(true);
		urgente.setFontName("Lato UI");
		
		for (Sheet s : wb) {

			int i = 12; // la riga a cui appare la prima occorrenza di urgenza
			int j = 3; // la riga dove inizia la barra di colore urgenza
			Cell c = s.getRow(i).getCell(startCol);
			while (s.getRow(i)!=null && s.getRow(i).getCell(startCol)!=null) { 
				// Purtroppo non basta che ci sia solo la seconda condizione, in quanto viene lanciato
				// un NPE se ci si riferisce a una riga inesistente
				c = s.getRow(i).getCell(startCol);
				System.out.println(c.getAddress());
				AreaRef urBar = new AreaRef(new CellRef(j, startCol - 2), new CellRef(j, startCol + 3));
				if (c.getCellType().equals(CellType.STRING)) {
					
					String urgenza = c.getStringCellValue();
					for (int index = startCol - 2; index <= urBar.getLastCellRef().getCol(); index++) {
						// <= perchè getCol ritorna numCol+1
						Cell urC = s.getRow(urBar.getFirstCellRef().getRow()).getCell(index);

						switch (urgenza.toLowerCase()) {
						case "bassa": {
							CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
									IndexedColors.SEA_GREEN.getIndex());
							CellStyle cs = wb.createCellStyle();
							cs.cloneStyleFrom(c.getCellStyle());
							cs.setFont(bassa);
							c.setCellStyle(cs);
							
						}
							break;

						case "media": {
							CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
									IndexedColors.LIGHT_ORANGE.getIndex());
							CellStyle cs = wb.createCellStyle();
							cs.cloneStyleFrom(c.getCellStyle());
							cs.setFont(media);
							c.setCellStyle(cs);
						}
							break;
							
						case "alta": {
							CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
									IndexedColors.RED1.getIndex());
							CellStyle cs = wb.createCellStyle();
							cs.cloneStyleFrom(c.getCellStyle());
							cs.setFont(alta);
							c.setCellStyle(cs);
						}
							break;
							

						case "urgente": {
							CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
									IndexedColors.BLACK.getIndex());
							CellStyle cs = wb.createCellStyle();
							cs.cloneStyleFrom(c.getCellStyle());
							cs.setFont(urgente);
							c.setCellStyle(cs);
						}
							break;
						}
					}
				}
				i += 12; // visto che la prossima occorrenza si troverà 11 celle sotto
				j += 12;
				
			} 
		}
		return wb;
	}

	/**
	 * Modifica il file "ticket_input.xlsx" applicando modifiche quali
	 * l'auto-allungamento delle colonne, i colori delle barre relativi al livello
	 * d'urgenza.
	 * 
	 * @throws IOException
	 */
	public static void adjustWorkbook() throws IOException {
		final InputStream in = new FileInputStream("src/main/resources/ticket/excel/ticket_input.xlsx");
		Workbook wb = WorkbookFactory.create(in);
		in.close();
		final OutputStream out = new FileOutputStream("src/main/resources/ticket/excel/ticket_input.xlsx");
		System.out.println("Aggiusto le colonne...");
		// wb = POIUtilities.fitColumns(wb, 0);
		System.out.println("Setto i colori delle urgenze...");
		wb = urgencyColor(wb, 2);
		wb = urgencyColor(wb, 9);
		wb = urgencyColor(wb, 16);
		wb = urgencyColor(wb, 23);
		wb = urgencyColor(wb, 30);
		wb.write(out);
		out.close();
	}

	/**
	 * Dal File f ricava un UIGridXmlObject, per ogni riga del quale genera oggetti
	 * Entry che verranno poi inseriti nel context.
	 * 
	 * @param context - Il context su cui si vuole agire.
	 * @param f       - Il file da cui ricavare l'UIGridXmlObject.
	 * @param nameVar - Il nome con cui inserire la variabile nel context.
	 * @return context riempito con la lista di Entry.
	 * @throws ParseException 
	 */
	public static Context process(Context context, File f, String nameVar) throws ParseException {
		List<Entry> master = new ArrayList<>();
		UIGridXmlObject u = new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(f, "UTF-8"));
		for (UIGridRow ur : u.getRows()) {
			Entry e = new Entry(ur);
			master.add(e);
		}
		context.putVar(nameVar, master);
		return context;
	}

	public static void main(String[] args) throws IOException, ParseException {
		final InputStream in = new FileInputStream("src/main/resources/ticket/excel/ticket_template.xlsx");
		final OutputStream out = new FileOutputStream("src/main/resources/ticket/excel/ticket_input.xlsx");
		final File daRilasciare = new File("src/main/resources/ticket/xml/ticket_da_rilasciare.xml");
		final File daTestare = new File("src/main/resources/ticket/xml/ticket_da_testare.xml");
		final File inCorso = new File("src/main/resources/ticket/xml/ticket_in_corso.xml");
		final File assegnato = new File("src/main/resources/ticket/xml/ticket_assegnato.xml");
		final File daAssegnare = new File("src/main/resources/ticket/xml/ticket_da_assegnare.xml");
		System.out.println("Inizio...");
		Context context = new Context();
		context = loadImages(context);
		context = process(context, daRilasciare, "daRilasciare");
		context = process(context, daTestare, "daTestare");
		context = process(context, inCorso, "inCorso");
		context = process(context, assegnato, "assegnato");
		context = process(context, daAssegnare, "daAssegnare");
		JxlsHelper.getInstance().processTemplate(in, out, context);
		in.close();
		out.close();

		adjustWorkbook();
		System.out.println("Fine.");
	}

}
