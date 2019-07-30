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

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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
	 * Modifica il file "ticket_input.xlsx" applicando modifiche quali
	 * l'auto-allungamento delle colonne, i colori delle barre relativi al livello
	 * d'urgenza.
	 * 
	 * @throws IOException
	 */
	public static void adjustWorkbook() throws IOException {
		final InputStream in = new FileInputStream("src/main/resources/ticket/excel/ticket_output.xlsx");
		Workbook wb = WorkbookFactory.create(in);
		in.close();
		final OutputStream out = new FileOutputStream("src/main/resources/ticket/excel/ticket_output.xlsx");
		System.out.println("Aggiusto le colonne...");
		// wb = POIUtilities.fitColumns(wb, 0);
		System.out.println("Setto i colori delle urgenze...");
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
		final OutputStream out = new FileOutputStream("src/main/resources/ticket/excel/ticket_output.xlsx");
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
