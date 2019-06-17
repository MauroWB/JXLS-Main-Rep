package com.smeup.excelfromxml;

import java.io.File;
import java.io.IOException;
import javax.xml.parsers.ParserConfigurationException;
import Smeup.smeui.uidatastructure.uigridxml.*;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

public class ReadXMLObject {

	public static Document parse() throws DocumentException {
		SAXReader reader = new SAXReader();
		Document document = reader.read(new File("src/main/resources/xml/example.xml"));
		return document;
	}

	public static void readColumnsAtRoW(int row, UIGridXmlObject u) {
		System.out.println("Riga: " + row);

		System.out.println("\n");
		for (int i = 0; i < u.getColumnsCount(); i++) {
			System.out.print(u.getValueForCell(row, i) + " | ");
		}
	}

	public static void readEverything(UIGridXmlObject u) {	
		for (int j = 0; j < u.getRowsCount(); j++) {
			System.out.print("| ");
			for (int i = 0; i < u.getColumnsCount(); i++) {
				System.out.print(u.getValueForCell(j, i) + " | ");
			}
			System.out.println("");
		}
	}
	
	public static String[][] copyValue (UIGridXmlObject u) {
		
		String[][] uig = new String[u.getRowsCount()][u.getColumnsCount()];
		
		for (int j = 0; j < u.getRowsCount(); j++) {
			for (int i = 0; i < u.getColumnsCount(); i++) {
				uig[j][i]=u.getValueForCell(j, i);
			}
		}
		return uig;
	}
	
	public static String[] copyCol (UIGridXmlObject u) {
		
		String [] uic = new String[u.getColumnsCount()];
		for(int i=0; i<u.getColumnsCount(); i++)
			uic[i] = u.getValueForCell(0, i);
		return uic;
	}
	
	static Logger logger = LoggerFactory.getLogger(ReadXMLObject.class);
	
	public static void main(String[] args)
			throws GridOperationException, DocumentException, SAXException, IOException, ParserConfigurationException {

		Document d = parse();
		// d =
		// UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/fromloocup.xml");

		UIGridXmlObject u = new UIGridXmlObject(d);
		System.out.println("Documento letto");
		readEverything(u);
		copyCol(u);
	}

}
