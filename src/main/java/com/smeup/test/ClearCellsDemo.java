package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.jxls.area.XlsArea;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.transform.poi.PoiTransformer;

public class ClearCellsDemo {

	public static void writeDown(Workbook wb, List<XlsArea> xlsAreaList) {
		XlsArea xlsArea = xlsAreaList.get(0);
		int actRow = xlsArea.getStartCellRef().getRow();
		int finRow = xlsArea.getAreaRef().getLastCellRef().getRow();
		System.out.println("Inizia a riga " + actRow + " e finisce a riga " + finRow);
		Sheet s = wb.getSheetAt(0);

		for (XlsArea x : xlsAreaList) {
			System.out.println(x.getAreaRef().getFirstCellRef());
			CellReference firstCellRef = new CellReference(x.getAreaRef().getFirstCellRef().toString(true));
			CellReference finalCellRef = new CellReference(x.getAreaRef().getLastCellRef().toString(true));
			// Test, sostituire il contenuto di tutte le celle dell'area
			for (int startingRow = firstCellRef.getRow(); startingRow <= finalCellRef.getRow(); startingRow++) {
				for (int startingColumn = firstCellRef.getCol(); startingColumn <= finalCellRef
						.getCol(); startingColumn++) {

					Cell c = s.getRow(startingRow).getCell(startingColumn);
					// System.out.println("Cella: " + c.getAddress());
					if (c != null && c.getCellType().equals(CellType.STRING)
							&& c.getStringCellValue().equalsIgnoreCase("clear"))
						x.clearCells();
				}
			}
			System.out.println("Passo alla prossima area...");
		}
		// Shiftare le aree?
		System.out.println("Job done");
	}

	/**
	 * Cancella le righe prive di valori.
	 * 
	 * @param workbook
	 * @param xlsAreaList
	 */
	public static void checkValues(Workbook wb, List<XlsArea> xlsAreaList) {
		Sheet s = wb.getSheetAt(0);
		boolean deleteRow = false;
		for (XlsArea x : xlsAreaList) {
			CellReference firstCellRef = new CellReference(x.getAreaRef().getFirstCellRef().toString(true));
			CellReference finalCellRef = new CellReference(x.getAreaRef().getLastCellRef().toString(true));

			for (int startingRow = firstCellRef.getRow(); startingRow <= finalCellRef.getRow(); startingRow++) {
				for (int startingColumn = firstCellRef.getCol() + 1; startingColumn <= finalCellRef
						.getCol(); startingColumn++) {
					Cell c = s.getRow(startingRow).getCell(startingColumn);
					if (c == null) {
						deleteRow = true;
					} else {
						deleteRow = false;
						break;
					}
				}
				if (deleteRow) {

					for (Cell del : s.getRow(startingRow)) {
						del.setCellType(CellType.BLANK);
						// TODO: Shiftare
						// s.shiftRows(s.getRow(startingRow).getRowNum()-1,
						// s.getRow(startingRow).getRowNum(), 1);
					}
				}
			}
		}
	}

	public static void main(String[] args) throws IOException {
		final InputStream in = new FileInputStream("src/main/resources/excel/clear/clear_template.xlsx");
		final OutputStream out = new FileOutputStream("src/main/resources/excel/clear/clear_output.xlsx");

		Workbook wb = WorkbookFactory.create(in);
		PoiTransformer transformer = PoiTransformer.createTransformer(wb);
		System.out.println(transformer.getCommentedCells().toString());
		AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer, false);
		List xlsAreaList = areaBuilder.build();
		System.out.println(xlsAreaList.get(0).toString());
		writeDown(wb, xlsAreaList);
		checkValues(wb, xlsAreaList);
		wb.write(out);
		wb.close();
		out.close();
		in.close();
	}

}
