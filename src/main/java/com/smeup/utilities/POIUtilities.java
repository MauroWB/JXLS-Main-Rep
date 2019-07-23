package com.smeup.utilities;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;


public class POIUtilities {
	public static Workbook fitColumns(Workbook wb) {
		for (Sheet s : wb)
			for (Row r : s)
				for (Cell c : r)
					s.autoSizeColumn(c.getColumnIndex());
		return wb;
	}
	
	/**
	 * Ridimensiona la lunghezza delle colonne, escludendo
	 * la colonna con il numero indicato.
	 * 
	 * @param wb - Workbook su cui agire
	 * @param excludeColumn - Il numero della colonna da escludere
	 * @return
	 */
	public static Workbook fitColumns(Workbook wb, int excludeColumn) {
		for (Sheet s : wb)
			for (Row r : s)
				for (Cell c : r)
					if (c.getColumnIndex()!=excludeColumn)
						s.autoSizeColumn(c.getColumnIndex());
						
		return wb;
	}
	
	/**
	 * Chiama il metodo com.smeup.utilities.POIUtilities.shiftRows(Sheet main, int first, int last, int n),
	 * dove come parametro "last" viene passata l'ultima riga del foglio main.
	 * 
	 * @param main - Foglio su cui applicare lo shift.
	 * @param first - Numero della riga di partenza.
	 * @param n - Numero di righe da shiftare (se n è negativo lo shift avverrà verso l'alto).
	 * @return Il foglio main con lo shift applicato.
	 */
	public static Sheet shiftRows(Sheet main, int first, int n) {
		return POIUtilities.shiftRows(main, first, main.getLastRowNum(), n);
	}
	
	/**
	 * Uguale al metodo org.apache.poi.ss.usermodel.Sheet.shiftRows(int startRow, int endRow, int n) ma funzionante,
	 * grazie alla patch ufficiale dal forum di Apache POI. Lo shift arriva sempre fino all'ultima riga del foglio.
	 * 
	 * @param main - Foglio su cui applicare lo shift.
	 * @param first - Numero della riga di partenza.
	 * @param last - Numero della riga d'arrivo.
	 * @param n - Numero di righe da shiftare (se n è negativo lo shift avverrà verso l'alto).
	 * @return Il foglio main con lo shift applicato.
	 */
	public static Sheet shiftRows(Sheet main, int first, int last, int n) {
		
		for (int i = 0; i < n; i ++) {
			main.createRow(last+1);
		}
		
		main.shiftRows(first, last, n);
		
		final int nFirstDstRow = first + n;
		final int nLastDstRow = last + n;
		
		for (int nRow = nFirstDstRow; nRow <= nLastDstRow; ++nRow) {
			final XSSFRow xrow = (XSSFRow) main.getRow(nRow);
			if (xrow != null) {
				String msg = "Row[rownum=" + xrow.getRowNum()
						+ "] contains cell(s) included in a multi-cell array formula. "
						+ "You cannot change part of an array.";
				for (Cell xc : xrow) {
					((XSSFCell) xc).updateCellReferencesForShifting(msg);
				}
			}
		}
		return main;
	}
}
