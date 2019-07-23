package com.smeup.commands;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jxls.area.Area;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.poi.PoiTransformer;

public class StyleCommand extends AbstractCommand {

	public static final String COMMAND_NAME = "style";
	String styleName;
	Area area;

	public String getName() {
		return "style";
	}

	public Size applyAt(CellRef cellRef, Context context) {
		String color = "WHITE";
		String fillType = "SOLID_FOREGROUND";
		// boolean border;
		boolean bold = false;
		boolean stop = false; // Se stop è true lo stile è stato trovato e non occorre proseguire i cicli.
		
		
		Size resultSize = area.applyAt(cellRef, context);
		PoiTransformer transformer = (PoiTransformer) area.getTransformer();
		Workbook wb = transformer.getWorkbook();

		Sheet main = wb.getSheet(cellRef.getSheetName());
		Sheet settings = wb.getSheet("Settings");
		
		if (resultSize.equals(Size.ZERO_SIZE) || styleName == null)
			return resultSize;
		
		// main = POIUtilities.shiftRows(main, cellRef.getRow(), resultSize.getHeight());
		// Non funziona
		
		styleName = styleName.trim().toUpperCase();
		
		CellStyle cs = wb.createCellStyle();
		
		int lastRow = cellRef.getRow() + resultSize.getHeight() - 1;
		int lastCol = cellRef.getCol() + resultSize.getWidth() - 1;
		// "-1" Perché i valori di sopra non sono 0 based
		for (int row = cellRef.getRow(); row <= lastRow; row++) {
			for (int col = cellRef.getCol(); col <= lastCol; col++) {

				if (main.getRow(row) == null)
					main.createRow(row);
				Cell c = main.getRow(row).getCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
				
				for (Row iR : settings) { 
					for (Cell iC : iR) {
							if (iC.getCellType().equals(CellType.STRING)
									&& styleName.equals(iC.getStringCellValue().trim().toUpperCase())) {

								// Controllo correttezza
								// TODO: Rifare questo codice che non si può vedere
								
								if (iR.getCell(iC.getColumnIndex() + 2).getCellType().equals(CellType.STRING))
									color = iR.getCell(iC.getColumnIndex() + 2).getStringCellValue();

								if (iR.getCell(iC.getColumnIndex() + 5).getCellType().equals(CellType.STRING))
									fillType = iR.getCell(iC.getColumnIndex() + 5).getStringCellValue();

								if (iR.getCell(iC.getColumnIndex() + 4).getCellType().equals(CellType.BOOLEAN)) 
									bold = iR.getCell(iC.getColumnIndex() + 4).getBooleanCellValue();
								System.out.println(bold);

								Font f = wb.createFont();
								f.setBold(bold);
								cs.setFont(f);

								// Gestione errori per evitare che l'input sbagliato dell'utente interrompa
								// l'esecuzione del programma.
								try {
									cs.setFillForegroundColor(IndexedColors.valueOf(color).index);
								} catch (Exception e) {
									System.out.println("Foreground Color non riconosciuto. Resetto a WHITE...");
									cs.setFillForegroundColor(IndexedColors.WHITE.index);
								}

								try {
									cs.setFillPattern(FillPatternType.valueOf(fillType));
								} catch (Exception e) {
									System.out.println("FillPatternType non riconosiuto. Resetto a SOLID_FOREGROUND...");
									cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
								}
								
								stop = true;
								break;
							}
						}
					if (stop)
						break;
				}
						

				// Nessun controllo visto che nessun attributo può essere null a questo punto
				
				c.setCellStyle(cs);
			}
		}
		return resultSize;
	}

	@Override
	public Command addArea(Area area) {
		super.addArea(area);
		this.area = area;
		return this;
	}

	public void setStyleName(String styleName) {
		this.styleName = styleName;
	}
}
