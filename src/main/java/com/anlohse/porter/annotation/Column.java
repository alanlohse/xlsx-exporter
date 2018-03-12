package com.anlohse.porter.annotation;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.ElementType.METHOD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

@Retention(RUNTIME)
@Target({ FIELD, METHOD })
public @interface Column {
	
	/**
	 * @return the index of field in the sheet. -1 means using the order of fields.
	 */
	int index() default -1;
	
	/**
	 * @return the header text of the column.
	 */
	String header() default "";

}
