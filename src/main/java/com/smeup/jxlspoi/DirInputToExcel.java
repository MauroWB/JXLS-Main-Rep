package com.smeup.jxlspoi;

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
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import com.smeup.test.SimpleGridObject;

import Smeup.smeui.uidatastructure.uigridxml.UIGridRow;
import Smeup.smeui.uiutilities.UIXmlUtilities;

/*
 * Printa n input su Excel
 *
 * Scorre tutti i file della directory, genera un nuovo 
 * oggetto s (un UIGridXMLObject praticamente) per ogni 
 * file, lo aggiunge al context e successivamente ne aggiunge
 * le singole colonne
 * 
 * Aggiunta funzionalità di auto lunghezza delle colonne
 * 
 */
public class DirInputToExcel {

	public static void autoWidth(Context context, InputStream in) throws IOException {
		Workbook wb = new XSSFWorkbook(in);
		OutputStream out = new FileOutputStream("src/main/resources/excel/dir_output.xlsx");

		Font ok = wb.createFont();
		ok.setColor(IndexedColors.GREEN.index);
		Font url = wb.createFont();
		url.setColor(IndexedColors.BLUE.index);

		for (Sheet s : wb)
			for (Row r : s)
				for (Cell c : r) {
					s.autoSizeColumn(c.getColumnIndex());
					if (c.getCellType() == CellType.STRING) {
						if (c.getStringCellValue().equals("C")) {
							CellStyle style = wb.createCellStyle();
							style.cloneStyleFrom(c.getCellStyle());
							style.setFont(ok);
							c.setCellStyle(style);
						}
						/*
						 * In Excel gli URL vengono già messi in blu, tuttavia in questo metodo la cosa
						 * è fatta per puro esercizio
						 */
						if (c.getCellType() == CellType.STRING && c.getStringCellValue().startsWith("http")) {
							CellStyle style = wb.createCellStyle();
							style.cloneStyleFrom(c.getCellStyle());
							style.setFont(url);
							c.setCellStyle(style);
							
						}
					}
				}
		wb.write(out);
		in.close();
		out.close();
		wb.close();
	}

	public static void fillContext(Context context, File dir) {
		// contatore per dare un nome a ciascuna variabile tipo "s1"
		int cont = 1;
		for (File f : dir.listFiles()) {
			SimpleGridObject s = new SimpleGridObject(UIXmlUtilities.buildDocumentFromXmlFile(f));
			s.setName("s" + cont);
			System.out.println(s.getName());
			// mette le colonne
			for (int i = 0; i < s.getColumnsCount(); i++) {
				List<Object> obj = new ArrayList<>();
				obj = Arrays.asList(s.getFormattedColumnValues(s.getColumnByIndex(i).getCod()));
				context.putVar(s.getName() + "_col" + (i + 1), obj);
				// mette nel context un array con tutte le colonne
				context.putVar(s.getName() + "_cols", Arrays.asList(s.getColumns()));
			}
			context.putVar(s.getName(), s.getTable());
			cont++;
		}
	}

	public static void main(String[] args) throws IOException {
		System.out.println("Inizio...");
		File dir = new File("src/main/resources/xml/xmltest");
		if (dir.isDirectory()) {
			InputStream in = new FileInputStream("src/main/resources/excel/dir_template.xlsx");
			OutputStream out = new FileOutputStream("src/main/resources/excel/dir_output.xlsx");
			Context context = new Context();
			fillContext(context, dir);
			JxlsHelper.getInstance().processTemplate(in, out, context);
			in.close();
			out.close();
			InputStream is = new FileInputStream("src/main/resources/excel/dir_output.xlsx");
			autoWidth(context, is);
		} else
			System.out.println("Il percorso indicato non è una directory.");
		System.out.println("Fine.");
	}

}
