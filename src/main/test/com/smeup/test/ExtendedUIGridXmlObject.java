package com.smeup.test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.dom4j.Document;

import Smeup.smec_s.utility.FileUtility;
import Smeup.smec_s.utility.LogManager;
import Smeup.smec_s.utility.SmeupFormatUtility;
import Smeup.smec_s.utility.SmeupFormulaSolver;
import Smeup.smec_s.utility.SmeupUtility;
import Smeup.smec_s.utility.StringUtility;
import Smeup.smeui.uidatastructure.uigridxml.UIGridColumn;
import Smeup.smeui.uidatastructure.uigridxml.UIGridXmlObject;
import Smeup.smeui.uimainmodule.MainApplication;
import Smeup.smeui.uimainmodule.MainGuiFrame;
import Smeup.smeui.uimainmodule.variables.UIVariableManager;

public class ExtendedUIGridXmlObject extends UIGridXmlObject {
	
	private ArrayList<List<Object>> table = new ArrayList<List<Object>>();
	private String name;

	public ExtendedUIGridXmlObject(Document document) {
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
	
	/*
	 * Metodi copiati al fine di garantire l'override di getValueForCell
	 * Non sembra funzionare
	 */
	private String prepareFormula(String aFormula, int aRowIndex, int aColumnIndex)
	{
		String vFormula = aFormula;
		if (vFormula.indexOf('[') != -1)
		{
			if ("[SUM]".equalsIgnoreCase(aFormula))
			{
				vFormula = "";
				UIGridColumn[] vColArr = getColumns();
				for (int i = 0; i < vColArr.length; i++)
				{
					UIGridColumn column = vColArr[i];
					if (aColumnIndex != i)
					{
						String vColOgg = column.getOgg();
						if (SmeupUtility.isNumericObject(vColOgg))
						{
							vFormula = vFormula + ("".equalsIgnoreCase(vFormula) ? "" : "+") + "[" + column.getCod() + "]";
						}
					}
				}
			}
			String vColName = null;
			while (true)
			{
				// Cerca i Nomi colonne e li sostituisce ai valori presenti nella riga per tali colonne al fine di risolvere la formula
				vColName = StringUtility.substringBetween(vFormula, "[", "]");
				if (StringUtility.isEmptyString(vColName))
				{
					break;
				}
				else
				{
					int vColIndex = getColumnIndex(vColName);
					String vPointedValue = "";
					if (vColIndex > -1)
					{
						vPointedValue = getValueForCell(aRowIndex, vColIndex);
					}
					vFormula = StringUtility.replace(vFormula, "[" + vColName + "]", "(" + vPointedValue + ")");
				}
			}
		}
		return vFormula;
	}
	
	private String getValueForCellFormula(String aFormula, int aRowIndex, int aColumnIndex)
	{
		// risolvo i riferimenti ai valori relativi alle colonne sulla riga
		if (MainApplication.DEBUG)
		{
			LogManager.log(MainGuiFrame.LOG_CODE, "Formula trovata: " + aFormula);
		}
		String vPreparedFormula = prepareFormula(aFormula, aRowIndex, aColumnIndex);
		if (MainApplication.DEBUG)
		{
			LogManager.log(MainGuiFrame.LOG_CODE, "Formula preparata: " + vPreparedFormula);
		}
		String vSolvedFormula = null;
		if (MainApplication.DEBUG)
		{
			LogManager.log(MainGuiFrame.LOG_CODE, "Richiedo risoluzione formula: " + vPreparedFormula);
		}
		vSolvedFormula = new SmeupFormulaSolver(new Character(getDecimalSeparator().charAt(0)), new Character(getThousandSeparator().charAt(0))).solveComplexExpression(vPreparedFormula);
		if (MainApplication.DEBUG)
		{
			LogManager.log(MainGuiFrame.LOG_CODE, "Ottengo formula risolta: " + vSolvedFormula);
		}
		return vSolvedFormula;
	}
	
	@Override
	public String getValueForCell(int aRowIndex, int aColumnIndex)
	{
		String vValueForCell = null;
		if (aColumnIndex < 0)
		{
			return vValueForCell;
		}
		UIGridColumn vColumn = getColumnByIndex(aColumnIndex);
		String vFormula = (vColumn != null) ? vColumn.getFormula() : null;
		// Se è una colonna formula, procedo al calcolo del valore
		if (vFormula != null && !"".equalsIgnoreCase(vFormula))
		{
			vValueForCell = this.getValueForCellFormula(vFormula, aRowIndex, aColumnIndex);
		}
		if (aRowIndex < this.getRowsCount() && aColumnIndex < this.getColumnsCount())
		{
			String[] vRowValues = this.getRows()[aRowIndex].getValues();
			if (vRowValues.length > aColumnIndex && (vValueForCell == null || "".equalsIgnoreCase(vValueForCell)))
			{
				vValueForCell = vRowValues[aColumnIndex];
				if (vColumn.isNumericColumn())
				{
					if ((vValueForCell == null) || ("".equalsIgnoreCase(vValueForCell.trim())))
					{
						vValueForCell = "0";
					}
					else if (vValueForCell.trim().endsWith("-"))
					{
						vValueForCell = "-" + vValueForCell.substring(0, vValueForCell.length() - 1);
					}
				}
			}
			else if (vColumn.isNumericColumn() && !vColumn.hasFormula())
			{
				try
				{
					if (vValueForCell == null)
					{
						vValueForCell = "0";
					}
					if (vValueForCell.trim().endsWith("-"))
					{
						vValueForCell = "-" + vValueForCell.substring(0, vValueForCell.length() - 1);
					}
					Double.parseDouble(vValueForCell);
				}
				catch (NumberFormatException e)
				{
					vValueForCell = "0";
					Smeup.smec_s.utility.LogManager.error(e);
				}
			}
			else if (!vColumn.hasFormula())
			{
				vValueForCell = "";
			}
			// Da togliere se non funzionante
			else if (vColumn.getOgg().equals("J4IMG")) {
				vValueForCell = "$SMEUP,IMG,"+vValueForCell+"$";
			}
		}
		return vValueForCell;
	}

	
	
	/**
	 * Alternativa di getFormattedValueForCell, fatta in modo che
	 * se qualora l'oggetto sia un'immagine, il risultato
	 * elaborato cominci con "$SMEUP" e finisca con "$".
	 */
	@SuppressWarnings("unused")
	public Object getFormattedValueForCellJXLS(int aRowIndex, String aColName) {
		Object vResult = null;
		String vFieldName = aColName;
		int vRowIndex = aRowIndex;
		int vColumnIndex = getColumnIndex(vFieldName);
		if (vColumnIndex != -1) {
			UIGridColumn vColumn = getColumnByCode(vFieldName);
			int vDeclaredLength = vColumn.getLength() + (vColumn.getNumDec() + 1); // Qui tengo conto del segno (se
																					// numero intero getNumDec= -1,
			// se decimale impongo numero dei decimali più uno per la virgola
			/**
			 * Qui vengono risolte eventuali formule presenti nella colonna
			 */
			String vValue = getValueForCell(vRowIndex, vColumnIndex);
			if (SmeupFormulaSolver.POSITIVE_INFINITY.equalsIgnoreCase(vValue)) {
				vResult = new Double(Double.POSITIVE_INFINITY);
				LogManager.log(MainGuiFrame.LOG_CODE, LogManager.T_INFO, vResult.toString());
			} else {
				if (vColumn != null && vColumn.isNumericColumn()) {
					vValue = StringUtility.replace(vValue, String.valueOf(this.getThousandSeparator()), "");
				}
				// OM 15/7/2011: Tolto per evitare taglio delle stringhe
				// if (vValue != null && vValue.length() > vDeclaredLength && vDeclaredLength >
				// 0)
				// {
				// vValue = vValue.substring(0, vDeclaredLength);
				// }
				vResult = vValue;
				UIGridColumn vCol = getColumnByIndex(vColumnIndex);
				if (vCol != null) {
					String vColType = vCol.getOgg();
					String vColLun = vCol.getLun();
					int vIntLun = 3;
					int vDecLun = 0;
					// E' un campo numerico se la lung è in formato NNN;NN o se il tipo è NR, in
					// quest'ultimo caso se contiene la "," ne deduco il formato
					// Intero,Decimale
					if (vColLun.indexOf(";") > -1) {
						vValue = StringUtility.replace(vValue, String.valueOf(this.getThousandSeparator()), "");
						String vIntLunString = vColLun.substring(0, vColLun.indexOf(";"));
						vIntLun = Integer.parseInt(vIntLunString);
						String vDecLunString = vColLun.substring(vColLun.indexOf(";") + 1);
						vDecLun = Integer.parseInt(vDecLunString);
					} else if (vValue != null && vValue.indexOf(this.getDecimalSeparator()) > -1) {
						vValue = StringUtility.replace(vValue, String.valueOf(this.getThousandSeparator()), "");
						String vPart1 = vValue.substring(0, vValue.indexOf(this.getDecimalSeparator()));
						String vPart2 = vValue.substring(vValue.indexOf(this.getDecimalSeparator()) + 1);
						// vValue= StringUtility.replace(vValue, String.valueOf(iDecimalSeparator), "");
						vIntLun = vPart1.length();
						vDecLun = vPart2.length();
					}
					if (SmeupUtility.isNumericObject(vColType)) {
						try {
							vResult = manageNumber(vValue, this.getThousandSeparator(), this.getDecimalSeparator());
						} catch (NumberFormatException exc) {
						}
					} else if (vColType.startsWith("D8")) {
						if (vColType.endsWith("*YYMD") || vColType.endsWith("*YMD") || vColType.endsWith("*DMYY")
								|| vColType.endsWith("*DMY")) {
							String vTip = vColType.substring(0, 2);
							String vPar = vColType.substring(2);
							if (vColType.endsWith("*YMD")) {
								vPar = "*DMY";
							}
							Date vJavaDate = SmeupFormatUtility.codeToDate(vTip, vPar, vValue);
							vResult = vJavaDate;
						}
					} else if (vColType.equalsIgnoreCase("J4IMG")) {
						String vImgPath = UIVariableManager.getVarValue("IMG:" + vValue);
						if (vImgPath != null && FileUtility.fileExists(vImgPath)) {
							// Se è una variabile immagine
							vResult = FileUtility.getFile(vImgPath);
							String value = "$SMEUP,IMG," + vResult.toString() + "$";
							vResult = value;
						} else if (vValue != null) {
							// Altrimenti la tratto come percorso
							String vSolvedPath = UIVariableManager.getVarValue(vValue);
							if (vSolvedPath != null
									&& FileUtility.fileExists(UIVariableManager.getVarValue(vSolvedPath))) {
								vResult = FileUtility.getFile(vSolvedPath);
								String value = "$SMEUP,IMG," + vResult.toString() + "$";
								vResult = value;
							}
						}
					} else if (vCol.isInstantColumn()) {
						if ("3".equalsIgnoreCase(vCol.getParametroSmeup())) {
							vValue = StringUtility.toSizedString((String) vValue, 4, false, '0');
						} else if ("4".equalsIgnoreCase(vCol.getParametroSmeup())) {
							vValue = StringUtility.toSizedString((String) vValue, 5, false, '0');
						} else// if("1".equalsIgnoreCase(vCol.getParametroSmeup()) ||
								// "1".equalsIgnoreCase(vCol.getParametroSmeup()))
						{
							vValue = StringUtility.toSizedString((String) vValue, 6, false, '0');
						}
						vResult = vValue;
					}
				}
			}
		}
		return vResult;
	}

	/**
	 * Alternativa al metodo getFormattedColumnValues, fatta di modo tale che ritorni
	 * un'ArrayList.
	 */
	public List<Object> getFormattedColumnValuesJXLS(String aColumnCode) {
		List<Object> vResult = new ArrayList<Object>();
		for (int i = 0; i < this.getRowsCount(); i++) {
			vResult.add(this.getFormattedValueForCellJXLS(i, aColumnCode));
		}
		return vResult;
	}
}
