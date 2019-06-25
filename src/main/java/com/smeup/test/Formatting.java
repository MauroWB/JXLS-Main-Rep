package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;

import org.dom4j.Document;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class Formatting {

	public static void main(String[] args) throws IOException {
		System.out.println("Inizio...");
		Document d = UIXmlUtilities
				.buildDocumentFromXmlFile("D:/Java/Workspace/ExcelFromXMLObject/src/main/resources/xml/fromloocup.xml");
		UIGridXmlObject uiGrid = new UIGridXmlObject(d);
		System.out.println("Documento letto");
		//SimpleGridObject gridC = new SimpleGridObject(uiGrid);

		System.out.println("Stampo la grid inalterata...");
		//gridC.printGrid();

		// Elaborazione template
		System.out.println("Procedo all'elaborazione del foglio Excel...");
		InputStream in = new FileInputStream("src/main/resources/excel/formatting_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/formatting_output.xlsx");
		Context context = new Context();
		context.putVar("headers", Arrays.asList(uiGrid.getColumnValues(0)));
		JxlsHelper.getInstance().processTemplate(in, out, context);

		System.out.println("Fine.");
		in.close();
		out.close();
	}

}
