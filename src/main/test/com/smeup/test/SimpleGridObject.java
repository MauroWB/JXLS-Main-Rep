package com.smeup.test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.dom4j.Document;

import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;

public class SimpleGridObject extends UIGridXmlObject {
	private ArrayList<List<Object>> table = new ArrayList<List<Object>>();
	private String name;

	public SimpleGridObject(Document document) {
		super(document);
		fillTable();
	}

	public void fillTable() {
		for (int i = 0; i < getColumnsCount(); i++)
			table.add(Arrays.asList(getFormattedColumnValues(getColumnByIndex(i).getCod())));
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getName() {
		return this.name;
	}

	public List<List<Object>> getTable() {
		return this.table;
	}

	public void printGrid() {
		for (List<Object> lObj : table) {
			System.out.println();
			for (Object obj : lObj)
				System.out.print(obj + " | ");
		}
	}

}
