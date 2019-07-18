package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
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

// Apparentemente è risaputo che shiftRows corrompa il file Excel
public class ClearCellsDemo {

	@SuppressWarnings("rawtypes")
	public static List removeAreas(Workbook wb, List<XlsArea> xlsAreaList) {

		List<XlsArea> toRemove = new ArrayList<XlsArea>();
		boolean deleteThisArea = false;
		Sheet s = wb.getSheetAt(0);

		for (XlsArea x : xlsAreaList) {
			deleteThisArea = false;
			System.out.println(x.getAreaRef().getFirstCellRef());
			CellReference firstCellRef = new CellReference(x.getAreaRef().getFirstCellRef().toString(true));
			CellReference finalCellRef = new CellReference(x.getAreaRef().getLastCellRef().toString(true));
			// Test, sostituire il contenuto di tutte le celle dell'area
			for (int currentRow = firstCellRef.getRow(); currentRow <= finalCellRef.getRow(); currentRow++) {
				for (int currentColumn = firstCellRef.getCol(); currentColumn <= finalCellRef
						.getCol(); currentColumn++) {

					Cell c = s.getRow(currentRow).getCell(currentColumn);
					if (c != null)
						System.out.println("Cella: " + c.getAddress()); // debug
					if (c != null && c.getCellType().equals(CellType.STRING)
							&& c.getStringCellValue().equalsIgnoreCase("clear")) {
						x.clearCells();
						deleteThisArea = true;
						break;
					}
				}
				if (deleteThisArea) {
					toRemove.add(x);
					break;
				}
			}

			System.out.println("Passo alla prossima area...");
		}
		// Rimozione aree da non considerare.
		xlsAreaList.removeAll(toRemove);
		System.out.println("Remove Areas completato.");
		return xlsAreaList;
	}

	/**
	 * Cancella le righe prive di valori.
	 * 
	 * @param workbook
	 * @param xlsAreaList
	 * @throws IOException
	 * @throws FileNotFoundException
	 */

	public static void checkValues(Workbook wb, List<XlsArea> xlsAreaList) throws FileNotFoundException, IOException {
		Sheet s = wb.getSheetAt(0);
		boolean deleteRow = false;
		for (XlsArea x : xlsAreaList) {
			CellReference firstCellRef = new CellReference(x.getAreaRef().getFirstCellRef().toString(true));
			System.out.println("Colonna della prima cella area: " + firstCellRef.getCol());
			CellReference lastCellRef = new CellReference(x.getAreaRef().getLastCellRef().toString(true));
			// Rircordarsi che CellRef.row è un numero in meno
			for (int currentRow = firstCellRef.getRow(); currentRow <= lastCellRef.getRow(); currentRow++) {
				deleteRow = false;
				System.out.println("Sono alla riga numero " + (currentRow + 1));
				for (int currentColumn = firstCellRef.getCol() + 1; currentColumn <= lastCellRef
						.getCol(); currentColumn++) { // + 1 perché vanno ignorate le celle "Valore:"

					if (s.getRow(currentRow).getCell(currentColumn) != null) {
						System.out.println(
								"Questa cella è ok: " + s.getRow(currentRow).getCell(currentColumn).getAddress());
						deleteRow = false;
						break;
					}
					deleteRow = true;
					System.out.println("Questa cella non è ok: " + new CellReference(currentRow, currentColumn));
					// oggetto CellReference per non dare la nullpointerexception
				}
				if (deleteRow) {

					System.out.println("----Voglio cancellare la riga n. " + (currentRow + 1));
					/*s.removeRow(s.getRow(currentRow));
					s.shiftRows(currentRow, s.getLastRowNum(), -1);*/

					System.out.println("pop");
					// Anche online si dice che questo metodo corrompa il foglio,
					// pure nei suoi più semplici utilizzi
				}
			}
		}

		System.out.println("Check values completato.");
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static void main(String[] args) throws IOException {
		final InputStream in = new FileInputStream("src/main/resources/excel/clear/clear_template.xlsx");
		final OutputStream out = new FileOutputStream("src/main/resources/excel/clear/clear_output.xlsx");

		//Workbook wb = WorkbookFactory.create(in);
		Workbook wb = WorkbookFactory.create(in);
		PoiTransformer transformer = PoiTransformer.createTransformer(wb);
		System.out.println(transformer.getCommentedCells().toString());
		AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer, false);
		List xlsAreaList = areaBuilder.build();
		System.out.println(xlsAreaList.get(0).toString());
		xlsAreaList = removeAreas(wb, xlsAreaList);
		checkValues(wb, xlsAreaList);
		Sheet s = wb.getSheetAt(0);
		s.shiftRows(3, 4, 1);
		wb.write(out);
		wb.close();
		out.close();
		in.close();
	}

}
