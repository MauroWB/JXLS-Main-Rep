package com.smeup.excelfromxml;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

/*
 * Usare la classe Grid per stampare i contenuti di
 * UIGridXmlObject
 */

public class UIGridToExcel {
	
	static Logger logger = LoggerFactory.getLogger(ReadXMLObject.class);
	
	//Sposta i contenuti di "u" in una lista
	public static ArrayList<ArrayList<String>> listify(UIGridXmlObject u) {
		ArrayList<ArrayList<String>> master = new ArrayList<ArrayList<String>>();
		for(int i=0; i<u.getRowsCount(); i++) {
			ArrayList<String> sublist = new ArrayList<String>();
			for(int j=0; j<u.getColumnsCount(); j++) {
				sublist.add(u.getValueForCell(i, j));
			}
			master.add(sublist);
		}
		return master;
	}
	
	public static void main(String[] args) throws DocumentException, IOException {
		System.out.println("Inizio...");
		Document d = UIXmlUtilities.buildDocumentFromXmlFile("D:/Java/Workspace/ExcelFromXMLObject/src/main/resources/xml/fromloocup.xml", "UTF-8");

		UIGridXmlObject u = new UIGridXmlObject(d);
		System.out.println("Documento letto");
		ArrayList<ArrayList<String>> master = new ArrayList<ArrayList<String>>();
		master = listify(u);
		
		FileInputStream is = new FileInputStream("src/main/resources/excel/grid_template.xlsx");
		OutputStream os = new FileOutputStream("src/main/resources/excel/grid_output.xlsx");
		Context context = new Context();
		context.putVar("master", master);
		System.out.println("Elaboro template...");
		JxlsHelper.getInstance().processTemplate(is, os, context);
		System.out.println("Fine.");
		is.close();
		os.close();
	}

}
