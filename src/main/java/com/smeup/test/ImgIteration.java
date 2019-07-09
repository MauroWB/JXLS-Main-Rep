package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jxls.common.Context;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class ImgIteration {
	
	public static void createTemp() throws IOException {
		InputStream in = new FileInputStream("src/main/resources/excel/img/imit_tempo.xlsx");
		Workbook wb = WorkbookFactory.create(in);
		
		
		wb.close();
		in.close();
	}
	
	public static void main(String[] args) throws IOException {
		//InputStream in = new FileInputStream("src/main/resources/excel/img/imit_template.xlsx");
		//OutputStream out = new FileOutputStream("src/main/resources/excel/img/init_output.xlsx");
		
		UIGridXmlObject uxo = new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/xml/img/test.xml"));
		Context context = new Context();
		int cont = 1;
		for (int i = 0; i < uxo.getColumnsCount(); i ++) {
			ArrayList<Object> sub = new ArrayList<>();
			sub.addAll(Arrays.asList(uxo.getFormattedColumnValues(uxo.getColumnByIndex(i).getCod())));
			context.putVar("col"+cont, sub);
			cont++;
		}
		
		createTemp();
		
		//out.close();
		//in.close();
	}

}
