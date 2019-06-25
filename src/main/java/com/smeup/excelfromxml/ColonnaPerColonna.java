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

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

/*
 * Fa la stampa, colonna per colonna, dei dati
 * Xml potendo così decidere quali colonne
 * mettere
 */

public class ColonnaPerColonna {

	public static void main(String[] args) throws IOException {

		// Creazione tabella
		SimpleGridObject s = new SimpleGridObject(
				new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/xml/fromloocup.xml")));

		// Elaborazione template
		System.out.println("Procedo all'elaborazione del foglio Excel...");
		InputStream in = new FileInputStream("src/main/resources/excel/cpc_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/cpc_output.xlsx");
		Context context = new Context();
		context.putVar("master", s.getTable());

		// Mette in context le n colonne contenenti i dati
		for (int i = 0; i < s.getU().getColumnsCount(); i++) {
			List<Object> obj = new ArrayList<>();
			obj = Arrays.asList(s.getU().getFormattedColumnValues(s.getU().getColumnByIndex(i).getCod()));
			context.putVar("col" + (i + 1), obj);
			System.out.println(context.getVar("col" + (i + 1)));
		}
		JxlsHelper.getInstance().processTemplate(in, out, context);
		in.close();
		out.close();
		System.out.println("Fine.");
	}

}
