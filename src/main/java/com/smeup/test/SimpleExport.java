package com.smeup.test;

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

import com.smeup.utilities.XmlWriter;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

//Creata per analisi con XmlWriter
public class SimpleExport {

	public static void execute() throws IOException {
		File dir = new File("src/main/resources/xml/generatorTest/test.xml");
		UIGridXmlObject u = new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(dir));
		List<List<Object>> master = new ArrayList<List<Object>>();
		for (int i = 0; i < u.getColumnsCount(); i++) {
			List<Object> sub = new ArrayList<Object>();
			sub.add(u.getColumnByIndex(i).getTxt());
			sub.addAll(Arrays.asList(u.getFormattedColumnValues(u.getColumnByIndex(i).getCod())));
			master.add(sub);
			System.out.println("Aggiunta colonna numero " + i);
		}
		InputStream is = new FileInputStream("src/main/resources/excel/export/export_input.xlsx");
		OutputStream os = new FileOutputStream("src/main/resources/excel/export/export_output.xlsx");

		Context context = new Context();
		context.putVar("master", master);
		JxlsHelper.getInstance().processTemplate(is, os, context);
		System.out.println("Fine");

		is.close();
		os.close();
	}

	public static void main(String[] args) throws IOException {
		XmlWriter.main(args);
		execute();
	}

}
