package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import com.smeup.utilities.POIUtilities;

import Smeup.smeui.uidatastructure.uigridxml.UIGridColumn;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class RichImg {

	public static void process(UIGridXmlObject uxo) throws IOException {

		InputStream inTemp = new FileInputStream("src/main/resources/excel/img/richimg_template.xlsx");
		Workbook wb = WorkbookFactory.create(inTemp);
		OutputStream outF = new FileOutputStream("src/main/resources/excel/img/richimg_tempo.xlsx");
		FileOutputStream out = new FileOutputStream("src/main/resources/excel/img/richimg_output.xlsx");
		Context context = new Context();
		int i = 0;
		for (UIGridColumn uc : uxo.getColumns()) {
			i++;
			context.putVar("col" + i, Arrays.asList(uxo.getFormattedColumnValues(uc.getCod())));
		}
		
		POIUtilities.autoIterate(wb, context);
		wb.write(outF);
		wb.close();
		outF.close();

		FileInputStream inF = new FileInputStream("src/main/resources/excel/img/richimg_tempo.xlsx");
		JxlsHelper.getInstance().processTemplate(inF, out, context);
		inF.close();
		out.close();
		Workbook wbF = WorkbookFactory.create(new FileInputStream("src/main/resources/excel/img/richimg_output.xlsx"));
		POIUtilities.insertImages(wbF);
		wbF.write(new FileOutputStream("src/main/resources/excel/img/richimg_final.xlsx"));
		System.out.println("Fine");
	}

	public static void main(String[] args) throws IOException {
		UIGridXmlObject uxo = new UIGridXmlObject(
				UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/xml/richimg.xml", "UTF-8"));
		process(uxo);
	}

}
