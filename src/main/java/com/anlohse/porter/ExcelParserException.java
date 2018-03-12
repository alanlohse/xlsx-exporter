package com.anlohse.porter;

public class ExcelParserException extends Exception {

	private static final long serialVersionUID = 1L;

	public ExcelParserException() {
	}

	public ExcelParserException(String message) {
		super(message);
	}

	public ExcelParserException(Throwable cause) {
		super(cause);
	}

	public ExcelParserException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelParserException(String message, Throwable cause, boolean enableSuppression,
			boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

}
