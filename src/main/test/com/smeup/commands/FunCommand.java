package com.smeup.commands;

import org.jxls.area.Area;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;

public class FunCommand extends AbstractCommand {
	
	public final static String COMMAND_NAME = "fun";
	
	String name;
	String fun;
	Area area;
	
	public String getName() {
		return "fun";
	}
	
	public void getDocumentFromFun() {
		
	}
	
	public Size applyAt(CellRef cellRef, Context context) {
		Size resultSize = area.applyAt(cellRef, context);
		// UIGridXmlObject u = new UIgridXmlObject(UIXmlUtilities.buildDocumentFromXmlFile(getDocumentFromFun()));
		// context.putVar(this.name, u);
		// u.setComment(this.name);
		return resultSize;
	}
	
	@Override
	public Command addArea(Area area) {
		super.addArea(area);
		this.area = area;
		return this;
	}
	
	public void setFun(String fun) {
		this.fun = fun;
	}
	
	public void setName(String name) {
		this.name = name;
	}
}
