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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import Smeup.smeui.uidatastructure.uigridxml.UIGridColumn;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uiutilities.UIXmlUtilities;

public class BetterExport {
	
	/**
	 * Genera un UIGridXmlObject dal primo file con estensione .xml
	 * trovato nella directory xmltest.
	 * 
	 * @return null se non esiste alcun file .xml, altrimenti l'UIGridXmlObject.
	 */
	public static UIGridXmlObject getUxoFromFile() {
		File dir = new File("src/main/resources/xml/xmltest");
		// Prende il primo file con estensione .xml
		File file = null; // Momentaneamente viene inizializzato come il primo elemento di dir
		for (File def : dir.listFiles())
			if (def.isFile() && def.getName().endsWith(".xml")) {
				file = def;
				break;
			}
		UIGridXmlObject u = new UIGridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(file, "UTF-8"));
		return u;
	}
	
	/**
	 * Inserisce alla prima riga del workbook le intestazioni di ciascuna colonna del file .xml.
	 * 
	 * @param wb - Il workbook su cui agire.
	 * @param u - L'UIGridXmlObject del file .xml
	 * @return il workbook modificato.
	 */
	public static Context generateHeaders(Context context, UIGridXmlObject u) 
	{
		List<String> headers = new ArrayList<>();
		for (UIGridColumn uc : u.getColumns()) 
		{
			String header = uc.getTxt()+"("+uc.getCod()+"|"+uc.getOgg()+"|"+uc.getLun()+")";
			headers.add(header);
		}
		context.putVar("headers", headers);
		return context;
	}
	
	/**
	 * Inserisce nel context i valori formattati di ciascuna colonna dell'UIGridXmlObject fornito.
	 * @param context - Il context su cui agire.
	 * @param u - L'UIGridXmlObject di cui si vogliono ricavare i dati.
	 * @return il context riempito.
	 */
	public static Context fillContext(Context context, UIGridXmlObject u) 
	{
		context = generateHeaders(context, u);
		int cont = 0;
		for (UIGridColumn uc : u.getColumns()) 
		{
			cont++;
			context.putVar("col"+cont, Arrays.asList(u.getFormattedColumnValues(uc.getCod())));
		}
		return context;
	}
	
	/**
	 * Scrive un workbook di template.
	 * 
	 * @param context - il context che verrà successivamente elaborato da Jxls.
	 * @return il context riempito.
	 * @throws IOException nel caso in cui non si possa accedere al path della destinazione.
	 */
	public static Context execute(Context context) throws IOException 
	{
		UIGridXmlObject u = getUxoFromFile();
		if (u == null) 
			return null;
		
		final OutputStream out = new FileOutputStream("src/main/resources/excel/export/export_temp.xlsx");
		context = fillContext(context, u);
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		CreationHelper factory = workbook.getCreationHelper();
		Drawing<?> drawing = sheet.createDrawingPatriarch();
		Row row = sheet.createRow(1);
		CellAddress last = new CellAddress(0,0);
		
		for (int i = 0; i < u.getColumnsCount(); i++) 
		{
			Cell c = row.createCell(i);
			c.setCellValue("${obj}");
			ClientAnchor anchor = factory.createClientAnchor();
			anchor.setRow1(c.getRowIndex()+1);
			anchor.setRow2(c.getRowIndex()+3);
			anchor.setCol1(c.getColumnIndex()+1);
			anchor.setCol2(c.getColumnIndex()+3);
			Comment comment = drawing.createCellComment(anchor);
			comment.setString(factory.createRichTextString("jx:each(lastCell='"+c.getAddress()
			+"' items='col"+(i+1)
			+"' var='obj' direction='DOWN')"));
			c.setCellComment(comment);
			System.out.println(comment);
			
			
			if (last.getRow() < c.getRowIndex()) 
			{
				int temp = last.getColumn();
				last = new CellAddress(c.getRowIndex(), temp);
			}
			if (last.getColumn() < c.getColumnIndex()) 
			{
				int temp = last.getRow();
				last = new CellAddress(temp, c.getColumnIndex());
			}
			
			
		}
		
		Cell origin = sheet.createRow(0).createCell(0);
		origin.setCellValue("${header}");
		ClientAnchor anchor = factory.createClientAnchor();
		Comment comment = drawing.createCellComment(anchor);
		XSSFRichTextString richTextString = (XSSFRichTextString) factory.createRichTextString("jx:area(lastCell='" + last + "')");
		richTextString.append("\njx:each(lastCell='A1' items='headers' var='header' direction='RIGHT')");
		comment.setString(richTextString);
		origin.setCellComment(comment);
		
		workbook.write(out);
		workbook.close();
		return context;
	}
	
	/**
	 * Fa il ridimensionamento automatico delle colonne. Scrive nello stesso file.
	 * @throws IOException
	 */
	public static void autoSize() throws IOException {
		final InputStream in = new FileInputStream("src/main/resources/excel/export/export_output.xlsx");
		Workbook workbook = WorkbookFactory.create(in);
		in.close();
		final OutputStream out = new FileOutputStream("src/main/resources/excel/export/export_output.xlsx");
		for (Sheet s : workbook)
			for (Row r : s)
				for (Cell c : r)
					s.autoSizeColumn(c.getColumnIndex());
		workbook.write(out);
		out.close();
		workbook.close();
	}

	public static void main(String[] args) throws IOException {

		Context context = new Context();
		try {
			context = execute(context);
		}
		catch (Exception e) {
			e.printStackTrace();
			return;
		}
		
		System.out.println("Inizio elaborazione jxls...");
		InputStream inTemp = new FileInputStream("src/main/resources/excel/export/export_temp.xlsx");
		OutputStream outF = new FileOutputStream("src/main/resources/excel/export/export_output.xlsx");
		JxlsHelper.getInstance().processTemplate(inTemp, outF, context);
		inTemp.close();
		outF.close();
		autoSize();
		System.out.println("Fine elaborazione.");
	}

}
