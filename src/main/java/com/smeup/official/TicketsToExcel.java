package com.smeup.official;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellUtil;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.jxls.util.Util;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;

import com.smeup.test.POIUtilities;

import Smeup.smeui.uidatastructure.uigridxml.UIGridRow;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class TicketsToExcel {
	final static String DESCRIZIONE = "ZÂ§AOBB"; // TODO: bisogna sempre mettere la A maiuscola? Risolvere

	/**
	 * Carica le immagini nel context come ByteArray
	 * 
	 * @param context
	 * @return
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

	// Attenzione, badare all'indicizzazione
	/**
	 * In base al contenuto della cella "Urgenza", cambia il colore della barra di
	 * ciascuna entry.
	 * 
	 * @param wb - Workbook su cui agire.
	 * @return
	 */
	public static Workbook urgencyColor(Workbook wb) {

		CellStyle bassa = wb.createCellStyle();
		bassa.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
		bassa.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		CellStyle media = wb.createCellStyle();
		media.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
		media.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		for (Sheet s : wb) {
			int i = 10; // la prima occorrenza di urgenza
			int j = 3; // dove inizia la barra di colore urgenza
			Cell c = s.getRow(i).getCell(1);
			while (c != null) {
				AreaRef urBar = new AreaRef(new CellRef(j, 0), new CellRef(j, 4));
				if (c.getCellType().equals(CellType.STRING)) {
					String urgenza = c.getStringCellValue();

					switch (urgenza.toLowerCase()) {
						case "bassa": {
							for (int index = 0; index <= urBar.getLastCellRef().getCol(); index++) { // <= perchè getCol
																										// ritorna numCol+1
								Cell urC = s.getRow(urBar.getFirstCellRef().getRow()).getCell(index);
								// urC.setCellStyle(bassa);
								CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
										IndexedColors.AQUA.getIndex());
								}
							}
							break;
						
						case "media": {
							for (int index = 0; index <= urBar.getLastCellRef().getCol(); index++) {
								Cell urC = s.getRow(urBar.getFirstCellRef().getRow()).getCell(index);
								CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
										IndexedColors.LIGHT_BLUE.getIndex());
								}
							}
							break;
							
						case "alta": {
							for (int index=0; index<=urBar.getLastCellRef().getCol(); index++) {
								Cell urC = s.getRow(urBar.getFirstCellRef().getRow()).getCell(index);
								CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
										IndexedColors.RED.getIndex());
							}
						}
						break;
						
						case "urgente": {
							for (int index=0; index<=urBar.getLastCellRef().getCol(); index++) {
								Cell urC = s.getRow(urBar.getFirstCellRef().getRow()).getCell(index);
								CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
										IndexedColors.DARK_RED.getIndex());
							}
						}
						break;
					}
				}

				i += 11; // visto che la prossima occorrenza si troverà 11 celle sotto
				j += 11;
				try {
					c = s.getRow(i).getCell(1);
				} catch (NullPointerException n) {
					break;
				}
			}
		}

		return wb;
	}
	
	public static Workbook condFormat(Workbook wb) {
		
		return wb;
	}

	/**
	 * Modifica il file "ticket_input.xlsx" applicando correzioni come
	 * l'allungamento automatico delle colonne
	 * 
	 * @throws IOException
	 */
	public static void adjustWorkbook() throws IOException {
		final InputStream in = new FileInputStream("src/main/resources/ticket/excel/ticket_input.xlsx");
		Workbook wb = WorkbookFactory.create(in);
		in.close();
		final OutputStream out = new FileOutputStream("src/main/resources/ticket/excel/ticket_input.xlsx");
		System.out.println("Aggiusto le colonne...");
		wb = POIUtilities.fitColumns(wb, 0);
		System.out.println("Setto i colori delle urgenze...");
		wb = urgencyColor(wb);
		System.out.println("Imposto le condizioni di formattazione...");
		wb = condFormat(wb);
		wb.write(out);
		out.close();
	}

	public static List<Entry> process(List<Entry> master) {
		final File f = new File("src/main/resources/ticket/xml/ticket_test.xml");
		UIGridXmlObject u = new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(f));
		for (UIGridRow ur : u.getRows()) {
			Entry e = new Entry(ur);
			master.add(e);
		}
		return master;
	}

	public static void main(String[] args) throws IOException {
		final InputStream in = new FileInputStream("src/main/resources/ticket/excel/ticket_template.xlsx");
		final OutputStream out = new FileOutputStream("src/main/resources/ticket/excel/ticket_input.xlsx");
		System.out.println("Inizio...");
		List<Entry> master = new ArrayList<>();
		master = process(master);
		
		Context context = new Context();
		context = loadImages(context);
		context.putVar("master", master);
		JxlsHelper.getInstance().processTemplate(in, out, context);

		in.close();
		out.close();

		adjustWorkbook();
		System.out.println("Fine.");
	}

}
