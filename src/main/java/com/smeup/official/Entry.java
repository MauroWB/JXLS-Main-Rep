package com.smeup.official;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import Smeup.smeui.uidatastructure.uigridxml.UIGridRow;

/**
 * Semplice classe che facilita il riferimento agli attributi di una
 * riga (Xml Loocup) in Jxls.
 * 
 * @author JohnSmith
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
		final DateFormat ideal = new SimpleDateFormat("dd/MM/yyyy");
		final DateFormat expected = new SimpleDateFormat("yyyyMMdd");
		Date date = expected.parse(ur.getValueForColumnCode(SCADENZA));
		//this.data = ideal.format(date);
		this.data = ideal.parse(ideal.format(date)); // da togliere in caso
		System.out.println(ur.getValueForColumnCode(SCADENZA)+ " | " +data);
		
	}
	
	public int getPercentuale() {
		return percentuale;
	}

	public Date getData() {
		return data;
	}

	public void setData(Date data) {
		this.data = data;
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
