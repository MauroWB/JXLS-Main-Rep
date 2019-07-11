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
		List<SimpleGridObject> l = new ArrayList<>();
		int cont = 1;
		for (File file : dir.listFiles()) {
			SimpleGridObject s = new SimpleGridObject(UIXmlUtilities.buildDocumentFromXmlFile(file));
			s.setName("s" + cont);
			l.add(s);
			cont++;
		}
		for (SimpleGridObject obj : l) {
			System.out.println("-------------");
			System.out.println("Tabella " + obj.getName()+"\n");
			obj.printGrid();
		}
	}

}
