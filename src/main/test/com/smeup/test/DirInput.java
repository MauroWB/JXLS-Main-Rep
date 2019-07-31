package com.smeup.test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import Smeup.smeui.uiutilities.UIXmlUtilities;

/*
 * Anziché dire esplicitamente quanti XML ci sono,
 * si specifica la directory in cui essi si trovano
 * e la si fa scorrere
 */

public class DirInput {

	public static void main(String[] args) {
		// Questa cartella contiene due xml
		File dir = new File("src/main/resources/xml/xmltest");
		List<ExtendedUIGridXmlObject> l = new ArrayList<>();
		int cont = 1;
		for (File file : dir.listFiles()) {
			ExtendedUIGridXmlObject s = new ExtendedUIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(file, "UTF-8"));
			s.setName("s" + cont);
			l.add(s);
			cont++;
		}
		for (ExtendedUIGridXmlObject obj : l) {
			System.out.println("-------------");
			System.out.println("Tabella " + obj.getName()+"\n");
			obj.printGrid();
		}
	}

}
