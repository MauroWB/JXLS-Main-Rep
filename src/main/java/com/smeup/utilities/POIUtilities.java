package com.smeup.utilities;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;


public class POIUtilities {
	public static Workbook fitColumns(Workbook wb) {
		for (Sheet s : wb)
			for (Row r : s)
				for (Cell c : r)
					s.autoSizeColumn(c.getColumnIndex());
		return wb;
	}
	
	/**
	 * In base al contenuto della cella "Urgenza", cambia il colore della barra di
	 * ciascuna entry.
	 * 
	 * @param wb - Workbook su cui agire.
	 * @return Workbook con colori relativi a ciascun campo urgenza.
	 */
	public static Workbook urgencyColor(Workbook wb, int startCol) {

		Font bassa = wb.createFont();
		bassa.setColor(IndexedColors.SEA_GREEN.getIndex());
		bassa.setBold(true);
		bassa.setFontName("Lato UI");
		
		Font media = wb.createFont();
		media.setColor(IndexedColors.LIGHT_ORANGE.getIndex());
		media.setBold(true);
		media.setFontName("Lato UI");
		
		Font alta = wb.createFont();
		alta.setColor(IndexedColors.RED.getIndex());
		alta.setBold(true);
		alta.setFontName("Lato UI");
		
		Font urgente = wb.createFont();
		urgente.setColor(IndexedColors.BLACK.getIndex());
		urgente.setBold(true);
		urgente.setFontName("Lato UI");
		
		for (Sheet s : wb) {

			int i = 12; // la riga a cui appare la prima occorrenza di urgenza
			int j = 3; // la riga dove inizia la barra di colore urgenza
			Cell c = s.getRow(i).getCell(startCol);
			while (s.getRow(i)!=null && s.getRow(i).getCell(startCol)!=null) { 
				// Purtroppo non basta che ci sia solo la seconda condizione, in quanto viene lanciato
				// un NPE se ci si riferisce a una riga inesistente
				c = s.getRow(i).getCell(startCol);
				System.out.println(c.getAddress());
				AreaRef urBar = new AreaRef(new CellRef(j, startCol - 2), new CellRef(j, startCol + 3));
				if (c.getCellType().equals(CellType.STRING)) {
					
					String urgenza = c.getStringCellValue();
					for (int index = startCol - 2; index <= urBar.getLastCellRef().getCol(); index++) {
						// <= perch� getCol ritorna numCol+1
						Cell urC = s.getRow(urBar.getFirstCellRef().getRow()).getCell(index);

						switch (urgenza.toLowerCase()) {
						case "bassa": {
							CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
									IndexedColors.SEA_GREEN.getIndex());
							CellStyle cs = wb.createCellStyle();
							cs.cloneStyleFrom(c.getCellStyle());
							cs.setFont(bassa);
							c.setCellStyle(cs);
							
						}
							break;

						case "media": {
							CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
									IndexedColors.LIGHT_ORANGE.getIndex());
							CellStyle cs = wb.createCellStyle();
							cs.cloneStyleFrom(c.getCellStyle());
							cs.setFont(media);
							c.setCellStyle(cs);
						}
							break;
							
						case "alta": {
							CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
									IndexedColors.RED1.getIndex());
							CellStyle cs = wb.createCellStyle();
							cs.cloneStyleFrom(c.getCellStyle());
							cs.setFont(alta);
							c.setCellStyle(cs);
						}
							break;
							

						case "urgente": {
							CellUtil.setCellStyleProperty(urC, CellUtil.FILL_FOREGROUND_COLOR,
									IndexedColors.BLACK.getIndex());
							CellStyle cs = wb.createCellStyle();
							cs.cloneStyleFrom(c.getCellStyle());
							cs.setFont(urgente);
							c.setCellStyle(cs);
						}
							break;
						}
					}
				}
				i += 12; // visto che la prossima occorrenza si trover� 11 celle sotto
				j += 12;
				
			} 
		}
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
	 * @param n - Numero di righe da shiftare (se n � negativo lo shift avverr� verso l'alto).
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
	 * @param n - Numero di righe da shiftare (se n � negativo lo shift avverr� verso l'alto).
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
