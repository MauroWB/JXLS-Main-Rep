package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import org.dom4j.Document;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class UIToExcelFilter {

	public static void main(String[] args) throws IOException {
		System.out.println("Inizio...");
		Document d = UIXmlUtilities
				.buildDocumentFromXmlFile("D:/Java/Workspace/ExcelFromXMLObject/src/main/resources/xml/fromloocup.xml", "UTF-8");
		UIGridXmlObject uiGrid = new UIGridXmlObject(d);
		System.out.println("Documento letto");
		SimpleGrid gridC = new SimpleGrid(uiGrid);

		System.out.println("Stampo la grid inalterata...");
		gridC.printGrid();

		// FrameFilter f = new FrameFilter("Filter");
		String filter = "Oggetto";
		System.out.println("\nOra stampo la nuova grid filtrata, senza i Txt " + filter);
		SimpleGrid fGrid = new SimpleGrid(uiGrid, filter);
		fGrid.printGrid();

		// Elaborazione template
		System.out.println("Procedo all'elaborazione del foglio Excel...");
		InputStream in = new FileInputStream("src/main/resources/excel/filter_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/filter_output.xlsx");
		Context context = new Context();
		context.putVar("grid", gridC);
		context.putVar("fGrid", fGrid);
		JxlsHelper.getInstance().processTemplate(in, out, context);
		in.close();
		out.close();
		System.out.println("Fine.");
	}

}
