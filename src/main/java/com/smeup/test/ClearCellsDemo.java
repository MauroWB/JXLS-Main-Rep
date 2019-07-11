package com.smeup.test;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.transform.poi.PoiTransformer;
//TODO: Da finire
public class ClearCellsDemo {
	
	public static void writeDown (Workbook wb, List<Area> xlsAreaList) {
		XlsArea xlsArea = (XlsArea) xlsAreaList.get(0);
		int actRow = xlsArea.getStartCellRef().getRow();
		int finRow = xlsArea.getAreaRef().getLastCellRef().getRow();
		System.out.println("Inizia a riga "+actRow+" e finisce a riga "+finRow);
		Sheet s = wb.getSheetAt(0);
		Row r = s.createRow(finRow+2);
		r.createCell(0).setCellValue("Test");
		
		for (Area x : xlsAreaList) {
			System.out.println(x.getAreaRef().getFirstCellRef());
			String g = x.getAreaRef().getFirstCellRef().toString(true);
			CellReference firstCellRef = new CellReference(g);
			CellReference finalCellRef = new CellReference(x.getAreaRef().getLastCellRef().toString(true));
			
			Row tr = s.getRow(firstCellRef.getRow());
			System.out.println("True that "+tr.getCell(firstCellRef.getCol()).getAddress());
			tr.getCell(firstCellRef.getCol()).setCellValue("L'area inizia qui");
			
			// Test, sostituire il contenuto di tutte le celle dell'area
			for (int startingRow = firstCellRef.getRow(); startingRow <= finalCellRef.getRow(); startingRow++) {
				for (int startingColumn = firstCellRef.getCol(); startingColumn <= finalCellRef.getCol(); startingColumn++) {
					//s.getRow(startingRow).getCell(startingColumn).setCellValue("Pulizia");
					if (s.getRow(startingRow).getCell(startingColumn).getStringCellValue().equalsIgnoreCase("clear"))
						((XlsArea) x).clearCells();
				}
			}
			
		}
		
		System.out.println("Job done");
	}

	public static void main(String[] args) throws IOException {
		final InputStream in = new FileInputStream("src/main/resources/excel/clear/clear_template.xlsx");
		final OutputStream out = new FileOutputStream("src/main/resources/excel/clear/clear_output.xlsx");
		
		Workbook wb = WorkbookFactory.create(in);
		for (Sheet s : wb)
			for (Row r : s)
				for (Cell c : r)
					System.out.println(c.getStringCellValue());
		PoiTransformer transformer = PoiTransformer.createTransformer(wb);
		System.out.println(transformer.getCommentedCells().toString());
		AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer, false);
		List<Area> xlsAreaList = areaBuilder.build();
		System.out.println(xlsAreaList.get(0).toString());
		writeDown(wb, xlsAreaList);
		wb.write(out);
		wb.close();
		out.close();
		in.close();
	}

}
