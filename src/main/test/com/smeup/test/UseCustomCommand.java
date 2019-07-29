package com.smeup.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;

import com.smeup.commands.ColorDemoCommand;
import com.smeup.commands.StyleCommand;
import com.smeup.commands.TemplateCommand;
import com.smeup.commands.TestCustomCommand;
import com.smeup.utilities.POIUtilities;

import Smeup.smeui.uiutilities.UIXmlUtilities;


public class UseCustomCommand {
	
	public static void main(String[] args) throws IOException {
		final InputStream in = new FileInputStream("src/main/resources/custom_commands/excel/command_template.xlsx");
		final OutputStream out = new FileOutputStream("src/main/resources/custom_commands/excel/command_output.xlsx");
		final Workbook wb = WorkbookFactory.create(new File("src/main/resources/custom_commands/excel/command_template.xlsx"));
		ExtendedUIGridXmlObject sgo = new ExtendedUIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile("src/main/resources/xml/richimg.xml"));
		
		
		Transformer transformer = TransformerFactory.createTransformer(in, out);
		AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
		XlsCommentAreaBuilder.addCommandMapping("smeuptest", TestCustomCommand.class);
		XlsCommentAreaBuilder.addCommandMapping(ColorDemoCommand.COMMAND_NAME, ColorDemoCommand.class);
		XlsCommentAreaBuilder.addCommandMapping(StyleCommand.COMMAND_NAME, StyleCommand.class);
		XlsCommentAreaBuilder.addCommandMapping(TemplateCommand.COMMAND_NAME, TemplateCommand.class);
		List<Area> xlsAreaList = areaBuilder.build();
		Context context = new Context();
		context = POIUtilities.cercaFun(context, wb.getSheet("FUN"));
		context.putVar("x", 5);
		context.putVar("sgo_test", sgo.getFormattedValueForCell(0, sgo.getColumnByIndex(0).getCod()));
		
		/*
		 * Se si aggiungono dei comandi custom, occorre applicare manualmente ciascuna XlsArea a cella e foglio
		 * a cui appartiene.
		 */
		for (Area xlsArea : xlsAreaList)
			xlsArea.applyAt(xlsArea.getStartCellRef(), context);
		
		// Corrispettivo di JxlsHelper.getInstance().processGridTemplate() ?
		transformer.write();
		in.close();
		out.close();
		System.out.println("Fine.");
	}

}
