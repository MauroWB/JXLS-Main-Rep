package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.dom4j.Document;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import Smeup.smeui.uiutilities.UIXmlUtilities;

/*
 * Stampa su Excel i dati da Xml
 * formattati
 */
public class ExtendedGridToExcel {

	public static void main(String[] args) throws IOException {
		System.out.println("Inizio...");
		Document d = UIXmlUtilities
				.buildDocumentFromXmlFile("D:/Java/Workspace/ExcelFromXMLObject/src/main/resources/xml/fromloocup.xml", "UTF-8");
		System.out.println("Documento letto");
		ExtendedUIGridXmlObject sgo = new ExtendedUIGridXmlObject(d);

		// Elaborazione template
		System.out.println("Procedo all'elaborazione del foglio Excel...");
		InputStream in = new FileInputStream("src/main/resources/excel/ex_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/ex_output.xlsx");
		Context context = new Context();
		context.putVar("sgo", sgo.getTable());
		JxlsHelper.getInstance().processTemplate(in, out, context);
		in.close();
		out.close();
		System.out.println("Fine.");
	}
}