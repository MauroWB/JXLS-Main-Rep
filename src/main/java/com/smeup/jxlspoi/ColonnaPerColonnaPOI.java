package com.smeup.jxlspoi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import com.smeup.test.ExtendedUIGridXmlObject;

import Smeup.smeui.uiutilities.UIXmlUtilities;

/*
 * Fa la stampa, colonna per colonna, dei dati
 * Xml potendo cos� decidere quali colonne
 * mettere.
 * 
 * Ora con due grid di due file xml
 * diversi
 * 
 * La versione ColonnaPerColonnaPOI contiene l'opzione
 * di invocare un metodo che, sfruttando le funzionalit�
 * di Apache POI, permette di autoformattare le colonne
 */

public class ColonnaPerColonnaPOI {

	/**
	 *  Mette nel context la grid e ciascuna colonna di cui � composta.
	 * @param context
	 * @param e
	 */
	public static void fillContext(Context context, ExtendedUIGridXmlObject e) {
		for (int i = 0; i < e.getColumnsCount(); i++) {
			List<Object> obj = new ArrayList<>();
			obj = Arrays.asList(e.getFormattedColumnValues(e.getColumnByIndex(i).getCod()));
			System.out.println("Aggiungo la colonna numero " + (i + 1) + " della grid " + e.getName() + "...");
			context.putVar(e.getName() + "_col" + (i + 1), obj);
			// qualcosa tipo "s_col1", "s1_col1"
		}
		context.putVar(e.getName(), e.getTable());
		context.putVar("uxo_" + e.getName(), e); // uxo inteso come UIGridXmlObject
		// Devo per forza mettere la lista di colonne come un Array, in quanto dentro
		// jxls non posso convertirlo
		context.putVar(e.getName() + "_columns", Arrays.asList(e.getColumns()));
		System.out.println("--Context riempito--\n");
	}

	
	/**
	 * Ridimensiona le colonne del workbook.
	 * @throws IOException
	 */
	public static void autoColumnWidth() throws IOException {
		System.out.println("Auto formattare colonne? (Y/N)");
		Scanner in = new Scanner(System.in);
		String choice = in.next().toLowerCase().trim();
		in.close();
		if (!(choice.equalsIgnoreCase("y")))
			return;

		System.out.println("Inizio elaborazione colonne...");
		InputStream win = new FileInputStream("src/main/resources/excel/cpc_output.xlsx");
		OutputStream wout = new FileOutputStream("src/main/resources/excel/cpcauto_output.xlsx");
		Workbook workbook = new XSSFWorkbook(win);
		Sheet sheet = workbook.getSheetAt(0);
		for (Row r : sheet) {
			for (Cell c : r)
			{
				sheet.autoSizeColumn(c.getColumnIndex());
			}
		}
		workbook.write(wout);
		workbook.close();
		wout.close();
		System.out.println("Fine elaborazione colonne.");
	}

	public static void main(String[] args) throws IOException {
		// Per debug
		long start = System.currentTimeMillis();

		// Creazione tabella
		ExtendedUIGridXmlObject e = new ExtendedUIGridXmlObject(
				(UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/xml/fromloocup.xml", "UTF-8")));
		e.setName("e");
		ExtendedUIGridXmlObject e1 = new ExtendedUIGridXmlObject(
				UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/xml/example.xml", "UTF-8"));
		e1.setName("e1");
		ExtendedUIGridXmlObject e2 = new ExtendedUIGridXmlObject(
				UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/xml/fromloocup2.xml", "UTF-8"));
		e2.setName("e2");

		// Elaborazione template
		System.out.println("Procedo all'elaborazione del foglio Excel...");
		InputStream in = new FileInputStream("src/main/resources/excel/cpc_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/cpc_output.xlsx");
		Context context = new Context();
		fillContext(context, e);
		fillContext(context, e1);
		fillContext(context, e2);

		JxlsHelper.getInstance().processTemplate(in, out, context);
		in.close();
		out.close();

		autoColumnWidth();
		// Chiusura stream
		long end = System.currentTimeMillis();
		System.out.println("Finito in " + (end - start) / 1000 + " secondi.");
	}

}
