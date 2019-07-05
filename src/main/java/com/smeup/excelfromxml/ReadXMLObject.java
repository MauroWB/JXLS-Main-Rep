package com.smeup.excelfromxml;

import java.io.IOException;

import javax.xml.parsers.ParserConfigurationException;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import Smeup.smeui.uidatastructure.uigridxml.GridOperationException;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class ReadXMLObject {

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
		Document d = UIXmlUtilities.buildDocumentFromXmlFile("D:/Java/Workspace/ExcelFromXMLObject/src/main/resources/xml/generatorTest/test.xml");
		
		System.out.println(d.getRootElement().selectNodes("//Riga").size());
		UIGridXmlObject u = new UIGridXmlObject(d);
		System.out.println("Documento letto");
		readEverything(u);
		copyCol(u);
	}

}
