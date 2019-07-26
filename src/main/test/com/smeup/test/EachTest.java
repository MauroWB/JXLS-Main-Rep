package com.smeup.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import Smeup.smeui.uidatastructure.uigridxml.UIGridRow;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class EachTest {

	public static void main(String[] args) throws IOException {
		final InputStream in = new FileInputStream("src/main/resources/test/each_template.xlsx");
		final OutputStream out = new FileOutputStream("src/main/resources/test/each_output.xlsx");
		final UIGridXmlObject u = new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(new File("src/main/resources/test/fromloocup.xml")));
		
		Context context = new Context();
		ArrayList<Object> arr = new ArrayList<>();
		for (int i = 0; i < 10; i++) 
			arr.add("T"+i);
		context.putVar("arr", arr);
		
		List<UIGridRow> sub = new ArrayList<>();
		for (UIGridRow ur : u.getRows()) 
			sub.add(ur);
		
		context.putVar("sub", sub);
		JxlsHelper.getInstance().processTemplate(in, out, context);
		in.close();
		out.close();
	}

}
