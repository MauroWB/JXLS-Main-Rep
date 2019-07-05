package com.smeup.utilities;

import java.io.File;
import java.util.Scanner;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class XmlWriter {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		String filePath = "D:\\Java\\Workspace\\ExcelFromXMLObject\\src\\main\\resources\\xml\\generatorTest\\test.xml";
		System.out.println("Inizio");

		try {
			Scanner s = new Scanner(System.in);
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dB = factory.newDocumentBuilder();
			Document document = dB.newDocument();

			Element root = document.createElement("UiSmeup");
			root.setAttribute("Testo", "  - ");
			document.appendChild(root);

			Element ser = document.createElement("Service");
			ser.setAttribute("DSep", ",");
			ser.setAttribute("Funzione", "F(EXB;B£SER_43B;EMIMAT) 1(;;) 2(;;) INPUT()");
			ser.setAttribute("IdFun", "0614154124983");
			ser.setAttribute("NumSes", "706949");
			ser.setAttribute("Servizio", "B£SER_43B");
			ser.setAttribute("TSep", ".");
			ser.setAttribute("Titolo1", "");
			ser.setAttribute("Titolo2", "");
			root.appendChild(ser);

			Element gri = document.createElement("Griglia");
			root.appendChild(gri);
			System.out.println("Numero colonne?");
			int cols = s.nextInt();

			// Crea 10 colonne
			for (int i = 0; i < cols; i++) {
				Element col = document.createElement("Colonna");
				col.setAttribute("Cod", "XXX" + i);
				col.setAttribute("IO", "O");
				col.setAttribute("Ogg", "NR");
				col.setAttribute("Txt", "Valore Campo " + i + "| XXX" + i);
				gri.appendChild(col);
			}

			Element rows = document.createElement("Righe");
			root.appendChild(rows);

			// Crea x righe a seconda dell'input
			System.out.println("Numero righe?");
			int input = s.nextInt();
			for (int i = 0; i < input; i++) {
				System.out.println("Creo riga numero " + i + "...");
				Element row = document.createElement("Riga");
				// Da fare per ogni colonna?
				//row.setAttribute("Fld", "00" + i + "|0" + i + "0" + "|TEST" + i + "|ESEMPIO " + i + "|TESTO " + i+ "|CELLA " + i + "| PROVA " + i);
				String fld = "";
				for (int j = 0; j < cols; j ++) {
					fld += "|Test"+j;
				}
				row.setAttribute("Fld", fld);
				rows.appendChild(row);
			}

			TransformerFactory tFactory = TransformerFactory.newInstance();
			Transformer transformer = tFactory.newTransformer();
			// Per fare l'indent, necessario quando si sta su numeri alti
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
			DOMSource domSource = new DOMSource(document);
			StreamResult out = new StreamResult(new File(filePath));

			transformer.transform(domSource, out);
			System.out.println("Fine");
			s.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
