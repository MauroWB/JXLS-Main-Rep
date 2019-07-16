package com.smeup.official;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Locale;
import java.util.Date;

import Smeup.smeui.uidatastructure.uigridxml.UIGridRow;

/**
 * Classe che facilita il riferimento agli attributi di una
 * riga (Xml Loocup) in Jxls
 * @author JohnSmith
 * 
 */

public class Entry {
	final String DESCRIZIONE = "Z§AOBB";
	final String UTENTE = "$$COEN_DE";
	final String URGENZA = "$$AIMP_DE";
	final String PERCENTUALE = "Z§NU01";
	final String SCADENZA = "Z§DADO";
	
	private String descrizione;
	private String utente;
	private String urgenza;
	private int percentuale;
	private Date data; 
	
	public Entry(UIGridRow ur) throws ParseException {
		this.descrizione=ur.getValueForColumnCode(DESCRIZIONE);
		this.utente = ur.getValueForColumnCode(UTENTE);
		this.urgenza = ur.getValueForColumnCode(URGENZA);
		this.percentuale = Integer.parseInt(ur.getValueForColumnCode(PERCENTUALE));
		DateFormat df = new SimpleDateFormat("YYYYMMDD");
		this.data = df.parse(ur.getValueForColumnCode(SCADENZA));
		System.out.println(this.data);
	}
	
	public int getPercentuale() {
		return percentuale;
	}

	public void setPercentuale(int percentuale) {
		this.percentuale = percentuale;
	}

	public String getDescrizione() {
		return descrizione;
	}
	public void setDescrizione(String descrizione) {
		this.descrizione = descrizione;
	}
	public String getUtente() {
		return utente;
	}
	public void setUtente(String utente) {
		this.utente = utente;
	}
	public String getUrgenza() {
		return urgenza;
	}
	public void setUrgenza(String urgenza) {
		this.urgenza = urgenza;
	}
}
