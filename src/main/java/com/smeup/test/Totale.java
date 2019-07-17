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
 * Pone una riga totale in fondo a determinate colonne
 */
public class Totale {

	public static void main(String[] args) throws IOException {
		System.out.println("Inizio...");
		Document d = UIXmlUtilities
				.buildDocumentFromXmlFile("D:/Java/Workspace/ExcelFromXMLObject/src/main/resources/xml/fromloocup.xml", "UTF-8");
		System.out.println("Documento letto");
		SimpleGridObject sgo = new SimpleGridObject(d);

		// Elaborazione template
		System.out.println("Procedo all'elaborazione del foglio Excel...");
		InputStream in = new FileInputStream("src/main/resources/excel/totale_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/totale_output.xlsx");
		Context context = new Context();
		context.putVar("sgo", sgo.getTable());
		JxlsHelper.getInstance().processTemplate(in, out, context);
		in.close();
		out.close();
		System.out.println("Fine.");
	}
}