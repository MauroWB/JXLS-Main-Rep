package com.smeup.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import Smeup.smeui.uidatastructure.uigridxml.UIGridColumn;
import Smeup.smeui.uidatastructure.uigridxml.UIGridRow;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class SpecialCharacters {

	public static void main(String[] args) throws IOException {
		final InputStream in = new FileInputStream(
				"src/main/resources/caratteri_speciali/excel/speciali_template.xlsx");
		final OutputStream out = new FileOutputStream(
				"src/main/resources/caratteri_speciali/excel/speciali_output.xlsx");
				
		
		File f = new File("src/main/resources/caratteri_speciali/xml/caratteri_speciali.xml");
		UIGridXmlObject u = new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(f, "UTF-8"));
		ArrayList<Object> master = new ArrayList<Object>();
		for (UIGridRow ur : u.getRows()) {
			ArrayList<Object> sub = new ArrayList<>();
			for (UIGridColumn uc : u.getColumns()) {
				sub.add(Arrays.asList(ur.getValueForColumnCode(uc.getCod())));
				System.out.println(sub);
			}
			master.add(sub);
		}
		
		Context context = new Context();
		context.putVar("caratteri", master);
		
		JxlsHelper.getInstance().processTemplate(in, out, context);
		out.close();
		in.close();
	}

}
