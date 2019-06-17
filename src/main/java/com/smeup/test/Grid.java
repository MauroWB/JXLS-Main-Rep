package com.smeup.test;

import java.util.ArrayList;
import java.util.List;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;

public class Grid {
	private List cell = new ArrayList();
	private List row = new ArrayList<List>();
	private UIGridXmlObject u = null;

	public Grid(UIGridXmlObject u) {
		this.u=u;
		this.fill();
	}
	
	public void fill() {
		for(int i=0; i<u.getRowsCount(); i++) {
			System.out.println("Leggo la riga n. "+i);
			for(int j=0; j<u.getColumnsCount(); j++) {
				cell.add(u.getValueForCell(i, j));
				System.out.print(cell.get(j)+" , ");
			}
			row.add(cell);
			cell.clear();
		}
	}
}
