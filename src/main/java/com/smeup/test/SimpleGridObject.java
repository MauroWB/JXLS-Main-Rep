package com.smeup.test;

import java.util.ArrayList;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;

public class SimpleGridObject {
	private ArrayList<ArrayList<Object>> table = new ArrayList<ArrayList<Object>>();
	private UIGridXmlObject u;
	private String name;

	public SimpleGridObject(UIGridXmlObject u) {
		this.u = u;
		this.fillTableByColumn();
	}

	// da cambiare, non si potrà mai riempire per righe
	public SimpleGridObject(UIGridXmlObject u, String filter) {
		this.u = u;
		System.out.println("Ok, riempio con il filtro " + filter);
		this.fillTableByColumn(filter);
		/*
		 * else this.fillTableByColumn(); } else this.fillTableByRow(); Ancora da
		 * elaborare
		 */
	}

	// Riempie la tabella mettendo in ciascuna lista gli elementi di ogni riga
	public void fillTableByRow() {
		table.clear();
		System.out.println("Filling table by rows...");
		for (int i = 0; i < u.getRowsCount(); i++) {
			ArrayList<Object> row = new ArrayList<Object>();
			for (int j = 0; j < u.getColumnsCount(); j++) {
				row.add(u.getFormattedValueForCell(i, u.getColumnByIndex(j).getCod()));
			}
			table.add(row);
		}
	}

	// Riempie la tabella mettendo in ciascuna lista gli elementi di ogni colonna
	public void fillTableByColumn() {
		System.out.println("Filling table by columns...");
		for (int i = 0; i < u.getColumnsCount(); i++) {
			ArrayList<Object> row = new ArrayList<Object>();
			for (int j = 0; j < u.getRowsCount(); j++) {
				row.add(u.getFormattedValueForCell(j, u.getColumnByIndex(i).getCod())); // non si può usare
																						// GetColumnValues perché
																						// ritorna un array non
																						// iterable
			}
			table.add(row);
		}
	}

	// Riempie la tabella escludendo le colonne che corrispondono al filtro
	public void fillTableByColumn(String filter) {
		System.out.println("-------------------------\n");
		for (int k = 0; k < u.getColumnsCount(); k++) {
			if (!(u.getColumnByIndex(k).getTxt().equals(filter))) {
				ArrayList<Object> temp = new ArrayList<Object>();
				for (int j = 0; j < u.getRowsCount(); j++) {
					temp.add(u.getValueForCell(j, k));
				}
				table.add(temp);
			} else if (u.getColumnByIndex(k).getTxt().equals(filter)) {
				System.out.println("Salto la colonna n. " + k);
				continue;
			}
		}
		System.out.println("Ora la tabella è lunga " + table.size() + " caselle.\n");
		System.out.println("-------------------------\n");
	}

	public void printGrid() {
		for (int i = 0; i < this.table.size(); i++) {
			System.out.print("Colonna n. " + i + " : ");
			for (int j = 0; j < this.table.get(i).size(); j++)
				System.out.print(this.table.get(i).get(j) + " | ");
			System.out.println();
		}

	}
	
	public void setName(String name) {
		this.name=name;
	}
	
	public String getName() {
		return this.name;
	}
	
	public UIGridXmlObject getU () {
		return this.u;
	}

	public ArrayList<ArrayList<Object>> getTable() {
		return this.table;
	}

}
