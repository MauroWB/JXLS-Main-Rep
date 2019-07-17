package com.smeup.utilities;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


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
}
