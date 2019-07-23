package com.smeup.commands;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jxls.area.Area;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.Util;

/**
 * Riempie le celle di una Jx Area con la stringa "Iterazione numero " + numero dell'iterazione. Utilizzo:
 * jx:smeupTest(lastCell="" condition="")
 *
 * @author JohnSmith
 *
 */

public class TestCustomCommand extends AbstractCommand {

	String condition;
	Area area;

	public String getName() {
		return "smeuptest";
	}

	// cellRef è dove starà il commento?

	public Size applyAt(CellRef cellRef, Context context) {
		Size resultSize = area.applyAt(cellRef, context);
		if (resultSize.equals(Size.ZERO_SIZE))
			return resultSize;
		PoiTransformer transformer = (PoiTransformer) area.getTransformer();
		Workbook wb = transformer.getWorkbook();
		Sheet s = wb.getSheet(cellRef.getSheetName());
		int endRow = cellRef.getRow() + resultSize.getHeight() - 1;
		int endCol = cellRef.getCol() + resultSize.getWidth() - 1;
		int cont = 0;
		if (condition != null && condition.trim().length() > 0) {
			boolean doIt = Util.isConditionTrue(getTransformationConfig().getExpressionEvaluator(), condition, context);
			if (doIt) {
				for (int row = cellRef.getRow(); row <= endRow; row++) {
					for (int col = cellRef.getCol(); col <= endCol; col++) {
						if (s.getRow(row) == null)
							s.createRow(row);
						if (s.getRow(row).getCell(col) == null)
							s.getRow(row).createCell(col);
						s.getRow(row).getCell(col).setCellValue("Iterazione numero " + cont);
						cont++;
					}
				}
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

	public void setCondition(String condition) {
		this.condition = condition;
	}
}
