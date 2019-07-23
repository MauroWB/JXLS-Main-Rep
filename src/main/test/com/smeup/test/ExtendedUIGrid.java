package com.smeup.test;

import java.util.ArrayList;
import java.util.List;

import org.dom4j.Document;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;

public class ExtendedUIGrid extends UIGridXmlObject {
	
	public ExtendedUIGrid(Document d) {
		super(d);
	}
	
	public List<Integer> getColumnValuesAsInt(int aColumnIndex) {

		List<Integer> columnValues = new ArrayList<Integer>(getRows().length);
		if (getColumns().length > aColumnIndex) {
			for (int rowID = 0; rowID < getRows().length; rowID++) {
				columnValues.add(rowID, Integer.parseInt(getRows()[rowID].getValues()[aColumnIndex]));
			}
			return columnValues;
		}
		return null;
	}
}
