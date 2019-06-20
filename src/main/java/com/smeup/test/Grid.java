package com.smeup.test;

import java.util.ArrayList;
import java.util.List;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;

public class Grid {
	private List<String> cell = new ArrayList<String>();
	private List<List<String>> row = new ArrayList<List<String>>();
	@SuppressWarnings("unused")
	private UIGridXmlObject u = null;

	public Grid(UIGridXmlObject u) {
		this.u=u;
		this.fill();
	}
	
	String[][] s  = {
			{
				"Col1","Col2","Col3"
			},
			{
				"Volvo", "G", "A"
			},
			{
				"12", "<", "3"
			}
	};
	
	public void fill() {
		for(int i=0; i<3; i++) {
			System.out.println("\nLeggo la riga n. "+i);
			for(int j=0; j<3; j++) {
				cell.add(s[i][j]);
				System.out.print(cell.get(j)+" | ");
			}
			row.add(cell);
			cell.clear();
		}
	}
}
