package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.dom4j.Document;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class NoGrid {

	public static void main(String[] args) throws IOException {
		Document d = UIXmlUtilities
				.buildDocumentFromXmlFile("D:/Java/Workspace/ExcelFromXMLObject/src/main/resources/xml/fromloocup.xml");
		UIGridXmlObject uiGrid = new UIGridXmlObject(d);
		
		List<List<String>> master = new ArrayList<List<String>>();
		for (int i=0; i<uiGrid.getColumnsCount(); i++) {
			List<String> a = Arrays.asList(uiGrid.getColumnValues(i));
			master.add(a);
		}
		//Iterable<String> iterable = Arrays.asList(cvalues);
		
		System.out.println("Inizio processazione...");
		InputStream in = new FileInputStream("src/main/resources/excel/nogrid_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/nogrid_output.xlsx");
		Context context = new Context();
		context.putVar("uiG", master);
		JxlsHelper.getInstance().processTemplate(in, out, context);
		in.close();
		out.close();
		System.out.println("Fine.");
	}

}
