package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import org.dom4j.Document;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class UIGridXmlObjectToExcel {

	public static void main(String[] args) throws IOException {
		System.out.println("Inizio...");
		Document d = UIXmlUtilities.buildDocumentFromXmlFile("D:/Java/Workspace/ExcelFromXMLObject/src/main/resources/xml/fromloocup.xml");

		UIGridXmlObject u = new UIGridXmlObject(d);
		System.out.println("Documento letto");
		System.out.println(u.toString());
		
		FileInputStream is = new FileInputStream("src/main/resources/excel/readloocup_template.xlsx");
		OutputStream os = new FileOutputStream("src/main/resources/excel/readloocup_output.xlsx");
		Context context = new Context();
		context.putVar("uigrid", u);
		System.out.println("Elaboro template...");
		JxlsHelper.getInstance().processTemplate(is, os, context);
		System.out.println("Fine.");
	}

}
