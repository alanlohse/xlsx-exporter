package com.anlohse.porter.types;

import java.io.Serializable;

public class Formula implements Serializable {

	private static final long serialVersionUID = 1L;
	public final String formula;

	public Formula(String formula) {
		this.formula = formula;
	}

	public String getFormula() {
		return formula;
	}
	
}
