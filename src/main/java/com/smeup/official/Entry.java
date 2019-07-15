package com.smeup.official;

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
	
	private String descrizione;
	private String utente;
	private String urgenza;
	private int percentuale;
	
	public Entry(UIGridRow ur) {
		this.descrizione=ur.getValueForColumnCode(DESCRIZIONE);
		this.utente = ur.getValueForColumnCode(UTENTE);
		this.urgenza = ur.getValueForColumnCode(URGENZA);
		this.percentuale = Integer.parseInt(ur.getValueForColumnCode(PERCENTUALE));
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
