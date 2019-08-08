package com.smeup.official;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.jxls.common.Context;
import org.omg.CORBA.portable.InputStream;
import org.omg.CORBA.portable.OutputStream;

import Smeup.smeui.uidatastructure.uigridxml.UIGridColumn;
import Smeup.smeui.uidatastructure.uigridxml.UIGridRow;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

/**
 * Inserisce nel context le fun, le righe, le colonne.
 * @author JohnSmith
 *
 */
public class FunContext {
	//TODO
	
	private static int funCount = 1;
	
	public static Context fillContext(UIGridXmlObject u) {
		
		Context context = new Context();
		File directory[] = new File("src/main/resources/xml/funcontext").listFiles();
		for (File f : directory) {
			if (f.getName().endsWith(".xml")) {
				
			}
		}
		
		List<Object> rows = null;
		for (int i = 0; i < u.getRowsCount(); i ++) {
			rows = new ArrayList<>();
			for (UIGridColumn uc : u.getColumns()) {
				rows.add(u.getFormattedValueForCell(i, uc.getCod()));
			}
			context.putVar("f"+funCount+"_row"+(i+1), rows);
		}
		
		return context;
	}
	
	public static void main (String args[]) throws IOException {
		
		InputStream in = null;
		OutputStream out = null;
		UIGridXmlObject u = null;
		Context context = fillContext(u);
		List<Object> rows = new ArrayList<>();
		
		try {
			
		}finally {
			in.close();
			out.close();
		}
	}
}
