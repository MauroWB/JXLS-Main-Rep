package com.smeup.test;

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
}
