package com.smeup.excelfromxml;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import com.smeup.test.SimpleGridObject;

import Smeup.smeui.uiutilities.UIXmlUtilities;

/*
 * Fa la stampa, colonna per colonna, dei dati
 * Xml potendo così decidere quali colonne
 * mettere.
 * 
 * Ora con le grid di due file xml
 * diversi
 *
 */

public class ColonnaPerColonna {

	/*
	 * Mette nel context la grid e ciascuna colonna di cui è composta
	 */
	public static void fillContext(Context context, SimpleGridObject s) {
		for (int i = 0; i < s.getColumnsCount(); i++) {
			List<Object> obj = new ArrayList<>();
			obj = Arrays.asList(s.getFormattedColumnValues(s.getColumnByIndex(i).getCod()));
			System.out.println("Aggiungo la colonna numero " + (i + 1) + " della grid " + s.getName() + "...");
			context.putVar(s.getName() + "_col" + (i + 1), obj);
			// qualcosa tipo "s_col1", "s1_col1"
		}
		context.putVar(s.getName(), Arrays.asList(s.getColumns()));
		System.out.println("--Context riempito--\n");
	}

	public static void main(String[] args) throws IOException {

		// Creazione tabella
		SimpleGridObject s = new SimpleGridObject(
				UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/xml/fromloocup.xml", "UTF-8"));
		s.setName("s");
		SimpleGridObject s1 = new SimpleGridObject(
				UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/xml/example.xml", "UTF-8"));
		s1.setName("s1");
		SimpleGridObject s2 = new SimpleGridObject(
				UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/xml/fromloocup2.xml", "UTF-8"));
		s2.setName("s2");

		// Elaborazione template
		System.out.println("Procedo all'elaborazione del foglio Excel...");
		InputStream in = new FileInputStream("src/main/resources/excel/cpc_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/cpc_output.xlsx");
		Context context = new Context();
		fillContext(context, s);
		fillContext(context, s1);
		fillContext(context, s2);

		JxlsHelper.getInstance().processTemplate(in, out, context);

		// Chiusura stream
		in.close();
		out.close();
		System.out.println("Fine.");
	}

}
