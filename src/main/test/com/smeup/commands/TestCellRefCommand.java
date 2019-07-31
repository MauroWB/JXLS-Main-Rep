package com.smeup.commands;

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

public class TestCellRefCommand extends AbstractCommand {
	
	public static final String COMMAND_NAME = "testCellRef";
	
	private Area area;
	
	@Override
	public String getName() {
		return "testCellRef";
	}

	@Override
	public Size applyAt(CellRef cellRef, Context context) {
		Size resultSize = area.applyAt(cellRef, context);
		PoiTransformer transformer = (PoiTransformer) area.getTransformer();
		Workbook wb = transformer.getWorkbook();
		Sheet s = wb.getSheet(cellRef.getSheetName());
		final int endRow = cellRef.getRow() + resultSize.getHeight()-1;
		final int endCol = cellRef.getCol() + resultSize.getWidth()-1;
		
		for (int row = cellRef.getRow(); row <= endRow; row ++)
			for (int col = cellRef.getCol(); col <= endCol; col++) {
				if (s.getRow(row) == null)
					s.createRow(row);
				Cell c = s.getRow(row).getCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
				CellUtil.setCellStyleProperty(c, CellUtil.BORDER_LEFT, true);
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
