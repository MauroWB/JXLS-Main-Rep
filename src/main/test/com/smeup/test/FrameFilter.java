package com.smeup.test;

import java.awt.BorderLayout;
import java.awt.Container;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JTextArea;
import javax.swing.JTextField;

/*
 * Impostare un filtro per rimuovere le colonne
 * con il txt indesiderato
 */

public class FrameFilter extends JFrame {

	private static final long serialVersionUID = 1L;
	public static JFrame frame = new JFrame("Input textbox");
	public static JButton inputButton = new JButton("Send");
	public static JTextArea textArea = new JTextArea("Type here!");
	public static JTextField textField = new JTextField();
	private String filter = "";

	public FrameFilter(String title) {
		
		frame.setLayout(new BorderLayout());
		textField.setEditable(false);
		

		Container c = frame.getContentPane();
		c.add(textField, BorderLayout.NORTH);
		c.add(textArea, BorderLayout.CENTER);
		c.add(inputButton, BorderLayout.SOUTH);

		inputButton.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				filter = textArea.getText();
				textArea.setText("");
				textField.setText("Fitler accepted");
				System.out.println(filter);
			}
		});
		frame.setSize(400, 500);
		frame.setVisible(true);
	}
	
	public String getFilter() {
		return this.filter;
	}
}
