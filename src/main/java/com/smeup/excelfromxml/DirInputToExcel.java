package com.smeup.excelfromxml;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import com.smeup.test.SimpleGridObject;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

/*
 * Printa n input su Excel
 *
 * Scorre tutti i file della directory, genera un
 * nuovo oggetto s (un UIGridXMLObject praticamente),
 * lo aggiunge al context e successivamente ne aggiunge
 * le singole colonne
 * 
 */
public class DirInputToExcel {

	public static void fillContext(Context context, File dir) {
		int cont = 1;

		for (File f : dir.listFiles()) {
			SimpleGridObject s = new SimpleGridObject(new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(f)));
			s.setName("s" + cont);
			System.out.println(s.getName());
			// mette le colonne
			for (int i = 0; i < s.getU().getColumnsCount(); i++) {
				List<Object> obj = new ArrayList<>();
				obj = Arrays.asList(s.getU().getFormattedColumnValues(s.getU().getColumnByIndex(i).getCod()));
				context.putVar(s.getName() + "_col" + (i + 1), obj);
			}
			context.putVar(s.getName(), s.getTable());
			cont++;
			s.printGrid();
		}
	}

	public static void main(String[] args) throws IOException {
		System.out.println("Inizio...");
		File dir = new File("src/main/resources/xml/xmltest");
		InputStream in = new FileInputStream("src/main/resources/excel/dir_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/dir_output.xlsx");
		Context context = new Context();
		fillContext(context, dir);
		JxlsHelper.getInstance().processTemplate(in, out, context);
		System.out.println(context.getVar("s1_col3"));
		in.close();
		out.close();
		out.flush();
		System.out.println("Fine.");
	}

}
