package com.smeup.commands;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.jxls.area.Area;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.poi.PoiTransformer;

/**
 * Crea dei bordi attorno all'area che parte dalla cella su cui è scritto il commento
 * a lastCell.
 * 
 * Utilizzo:
 * jx:template(lastCell="")
 * 
 * @author JohnSmith
 *
 */
public class TemplateCommand extends AbstractCommand {

	final static public String COMMAND_NAME = "template";
	private String templateName;
	private Area area;

	@Override
	public String getName() {
		return "template";
	}

	public String getTemplateName() {
		return templateName;
	}

	public void setTemplateName(String templateName) {
		this.templateName = templateName;
	}

	@Override
	public Size applyAt(CellRef cellRef, Context context) {
		
		Size resultSize = area.applyAt(cellRef, context);
		PoiTransformer transformer = (PoiTransformer) area.getTransformer();
		Workbook wb = transformer.getWorkbook();
		Sheet s = wb.getSheet(cellRef.getSheetName());
		final int lastRow = cellRef.getRow() + resultSize.getHeight() - 1;
		final int lastCol = cellRef.getCol() + resultSize.getWidth() - 1;
		int startRow = cellRef.getRow();
		final int startCol = cellRef.getCol();

		for (int row = cellRef.getRow(); row <= lastRow; row++)
			for (int col = cellRef.getCol(); col <= lastCol; col++) {
				
				if (s.getRow(row) == null)
					s.createRow(row);
				Cell c = s.getRow(row).getCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
				
				if (row == startRow)
					CellUtil.setCellStyleProperty(c, CellUtil.BORDER_TOP, BorderStyle.THICK);

				if (row == lastRow)
					CellUtil.setCellStyleProperty(c, CellUtil.BORDER_BOTTOM, BorderStyle.THICK);

				if (col == startCol)
					CellUtil.setCellStyleProperty(c, CellUtil.BORDER_LEFT, BorderStyle.THICK);

				if (col == lastCol)
					CellUtil.setCellStyleProperty(c, CellUtil.BORDER_RIGHT, BorderStyle.THICK);

			}
		return resultSize;
	}

	@Override
	public Command addArea(Area area) {
		super.addArea(area);
		this.area = area;
		return this;
	}

}
