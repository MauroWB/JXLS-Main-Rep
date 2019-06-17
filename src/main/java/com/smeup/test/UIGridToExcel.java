package com.smeup.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import org.jxls.transform.Transformer;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.io.SAXReader;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.smeup.excelfromxml.ReadXMLObject;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;

/*
 * Usare la classe Grid per stampare i contenuti di
 * UIGridXmlObject
 */

public class UIGridToExcel {
	
	static Logger logger = LoggerFactory.getLogger(ReadXMLObject.class);
	
	public static Document parse() throws DocumentException {
		SAXReader reader = new SAXReader();
		Document document = reader.read(new File("src/main/resources/xml/example.xml"));
		return document;
	}
	
	public static void main(String[] args) throws DocumentException, IOException {
		Document d = parse();

		UIGridXmlObject u = new UIGridXmlObject(d);
		System.out.println("Documento letto");
		Grid g = new Grid(u);
		
		FileInputStream is = new FileInputStream("src/main/resources/excel/grid_template.xlsx");
		OutputStream os = new FileOutputStream("src/main/resources/excel/grid_output.xlsx");
		Context context = new Context();
		context.putVar("grid", g);
		JxlsHelper.getInstance().processTemplate(is, os, context);
	}

}
