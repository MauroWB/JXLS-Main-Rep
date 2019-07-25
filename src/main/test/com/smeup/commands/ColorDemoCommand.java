package com.smeup.commands;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
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

/**
 * <p>Colora un'area a seconda del tipo di fill che si ha definito (non case sensitive):
 * <ul><li>FULL - Riempie tutta l'area.</li>
 * <li>CONTOUR - Riempie solo i contorni.</li>
 * </ul></p>
 * 
 * 
 * <b>Utilizzo:</b>
 * <br>
 * <p>jx:color(lastCell="" fillType="FULL")</p>
 * 
 * @author <p>JohnSmith</p>
 *
 */
public class ColorDemoCommand extends AbstractCommand {
	
	public final static String COMMAND_NAME = "color";
	private Area area;
	private String fillType;

	public String getName() {
		return "color";
	}

	public Size applyAt(CellRef cellRef, Context context) {
		
		Size resultSize = area.applyAt(cellRef, context);
		PoiTransformer transformer = (PoiTransformer) area.getTransformer();
		Workbook wb = transformer.getWorkbook();
		CellStyle testStyle = wb.createCellStyle();
		testStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
		testStyle.setFillPattern(FillPatternType.BRICKS);
		Sheet s = wb.getSheet(cellRef.getSheetName());
		if (resultSize.equals(Size.ZERO_SIZE))
			return resultSize;
		int endRow = cellRef.getRow() + resultSize.getHeight()-1;
		int endCol = cellRef.getCol() + resultSize.getWidth()-1;
		// Quindi se nel commento non è stato inserito il parametro "fillType"
		if (fillType == null)
			this.fillType = "FULL";
		for (int row = cellRef.getRow(); row <= endRow; row++)
			for (int col = cellRef.getCol(); col <= endCol; col++) {
				if (s.getRow(row) == null)
					s.createRow(row);
				Cell c = s.getRow(row).getCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
				switch (fillType.toLowerCase()) {

				case "full":
					c.setCellStyle(testStyle);
					break;

				case "contour": {
					if (row == cellRef.getRow() || row == endRow) {
						c.setCellStyle(testStyle);
						break;
					}
					if (col == cellRef.getCol() || col == endCol)
						c.setCellStyle(testStyle);
				}
					break;

				default:
					break;

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

	public void setFillType(String fillType) {
		this.fillType = fillType;
	}

}
