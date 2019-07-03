package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class AutoMarkUp {

	public static void read() throws IOException {
		InputStream in = new FileInputStream("src/main/resources/excel/auto/automu_template.xlsx");
		OutputStream out = new FileOutputStream("src/main/resources/excel/auto/automu_temp.xlsx");

		Workbook wb = WorkbookFactory.create(in);
		for (Sheet s : wb)
			for (Row r : s)
				for (Cell c : r) {
					if (c.getCellType() == CellType.STRING) {
						if (c.getStringCellValue().startsWith("<PRINT") && c.getStringCellValue().endsWith(">")) {
							/*
							 * Togli i < >, splitti per spazio, splitti per underscore - Se è null abbiamo
							 * un solo oggetto (tipo tabella) - Se lo split ritorna qualcosa, prendi
							 * l'elemento in pos.1, ne ricavi il numero se c'è scritto col. Sennò lasci
							 * stare tutto
							 */
							String elab = c.getStringCellValue();
							elab = elab.replace("<PRINT ", "");
							elab = elab.replaceAll(">", "");
							String[] res = elab.split("_");
							if (res[1].startsWith("col")) {
								// è ok e si può inserire nel context
								StringBuffer b = new StringBuffer();
								String[] fin = res[1].split("");
								// scorre alla ricerca di un numero colonna
								for (int i = 0; i < fin.length; i++) {
									if (Character.isDigit(fin[i].charAt(0)))
										b.append(fin[i]);
								}
								// TODO Metti nel context la colonnna all'indice Integer.parseInt(b)
							}
						}
					}
				}

		wb.write(out);
		out.close();
		wb.close();
		in.close();
	}
	
	
	
	public static void main(String[] args) throws IOException {
		read();
		// TODO: execute();
	}

}
