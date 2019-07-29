package com.smeup.excelfromxml;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.dom4j.Document;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import com.smeup.test.ExtendedUIGridXmlObject;

import Smeup.smeui.uiutilities.UIXmlUtilities;

public class UIGridXmlObjectToExcel {
	public static void main(String[] args) throws IOException {
		System.out.println("Inizio...");
		Document d = UIXmlUtilities
				.buildDocumentFromXmlFile("D:/Java/Workspace/ExcelFromXMLObject/src/main/resources/xml/fromloocup2.xml", "UTF-8");
		System.out.println("Documento letto");
		ExtendedUIGridXmlObject grid = new ExtendedUIGridXmlObject(d);

		FileInputStream is = new FileInputStream("src/main/resources/excel/readloocup_template.xlsx");
		OutputStream os = new FileOutputStream("src/main/resources/excel/readloocup_output.xlsx");
		Context context = new Context();
		context.putVar("uiGrid", grid.getTable());
		for (List<Object> o : grid.getTable()) {
			for (Object obj : o)
				System.out.println(obj);
		}
		context.putVar("grid", grid);
		System.out.println("Elaboro template...");
		JxlsHelper.getInstance().processTemplate(is, os, context);
		System.out.println("Fine.");
		is.close();
		os.close();
	}

}
