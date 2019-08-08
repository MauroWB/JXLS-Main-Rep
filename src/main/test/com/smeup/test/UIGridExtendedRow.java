package com.smeup.test;

import java.util.ArrayList;
import java.util.List;

import org.dom4j.Element;

import Smeup.smeui.uidatastructure.uigridxml.UIGridColumn;
import Smeup.smeui.uidatastructure.uigridxml.UIGridRow;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;

public class UIGridExtendedRow extends UIGridRow {

	private int iIndex;
	private UIGridXmlObject iUxo;
	
	public UIGridExtendedRow(UIGridRow aUIGridRow) 
	{
		super(aUIGridRow.getTipo(), aUIGridRow.getParametro(), aUIGridRow.getCodice(), aUIGridRow.getValues());
	}
	
	public Object getFormattedCellValue(String aColumnCod)
	{
		Object vResult = null;
		vResult = iUxo.getFormattedValueForCell(this.iIndex, aColumnCod);
		return vResult;
	}
	
	public List<Object> getFormattedCellValues()
	{
		List<Object> vResults = new ArrayList<Object>();
		
		for (UIGridColumn vColumn : iUxo.getColumns())
		{
			vResults.add(iUxo.getFormattedValueForCell(iIndex, vColumn.getCod()));
		}
		
		return vResults;
	}
	
	public void setIndex(int aIndex) 
	{
		this.iIndex = aIndex;
	}
	
	public void setUIGridXmlObject (UIGridXmlObject aUxo)
	{
		this.iUxo = aUxo;
	}
	
}
