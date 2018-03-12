package com.anlohse.porter.internal;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.util.Date;
import java.util.Locale;

import com.anlohse.porter.ExcelParserException;

public class ConvertOps {

	public static Object convert(Object value, Class<?> type, Locale locale) throws ExcelParserException {
		if (value == null && type.isPrimitive()) return primitiveNull(type);
		if (value == null) return value;
		if (value.getClass() == type || type.isAssignableFrom(value.getClass()))
			return value;
		if (value.getClass() == String.class && isNumber(type))
			return convertStringToNumber((String) value, type);
		if (value instanceof Number && isNumber(type))
			return convertNumberToNumber((Number)value, type);
		if (value.getClass() == String.class && (Boolean.class == type || boolean.class == type))
			return convertStringToBoolean((String) value);
		if (value.getClass() == String.class && Date.class == type)
			return convertStringToDate((String) value, locale);
		if (value instanceof Number && Date.class == type)
			return new Date(((Number) value).longValue());
		if (value instanceof Number && type == String.class)
			return NumberFormat.getNumberInstance().format((Number)value);
		if (value instanceof Date && type == String.class)
			return DateFormat.getDateInstance().format((Date)value);
		if (type == String.class)
			return value.toString();
		return value;
	}

	private static boolean isNumber(Class<?> type) {
		return type == int.class || type == double.class || type == long.class || type == float.class || type == short.class || type == byte.class
				 || Number.class.isAssignableFrom(type);
	}

	private static Object convertStringToDate(String value, Locale locale) throws ExcelParserException {
		try {
			return DateFormat.getDateInstance(DateFormat.SHORT, locale).parse(value);
		} catch (ParseException e) {
			try {
				return DateFormat.getDateTimeInstance(DateFormat.SHORT, DateFormat.MEDIUM, locale).parse(value);
			} catch (ParseException e1) {
				try {
					return DateFormat.getTimeInstance(DateFormat.MEDIUM, locale).parse(value);
				} catch (ParseException e2) {
					throw new ExcelParserException("Data inválida.");
				}
			}
		}
	}

	private static Object primitiveNull(Class<?> type) throws ExcelParserException {
		if (type == double.class)
			return (double)0d;
		if (type == int.class)
			return (int)0;
		if (type == long.class)
			return (long)0L;
		if (type == float.class)
			return (float)0f;
		if (type == short.class)
			return (short)0;
		if (type == byte.class)
			return (byte)0;
		if (type == boolean.class)
			return false;
		if (type == char.class)
			return (char)0;
		if (type == BigDecimal.class)
			return BigDecimal.ZERO;
		if (type == BigInteger.class)
			return BigInteger.ZERO;
		throw new ExcelParserException("Tipo numérico desconhecido.");
	}

	private static Object convertStringToBoolean(String value) {
		if (value.isEmpty()) return false;
		if (value.toLowerCase().equals("true")) return true;
		if (value.toLowerCase().equals("false")) return false;
		if (value.toLowerCase().equals("s")) return true;
		if (value.toLowerCase().equals("n")) return false;
		if (value.toLowerCase().equals("v")) return true;
		if (value.toLowerCase().equals("f")) return false;
		return true;
	}

	private static Object convertNumberToNumber(Number value, Class<?> type) throws ExcelParserException {
		if (type == double.class || type == Double.class)
			return value.doubleValue();
		if (type == int.class || type == Integer.class)
			return value.intValue();
		if (type == long.class || type == Long.class)
			return value.longValue();
		if (type == float.class || type == Float.class)
			return value.floatValue();
		if (type == short.class || type == Short.class)
			return value.shortValue();
		if (type == byte.class || type == Byte.class)
			return value.byteValue();
		if (type == BigDecimal.class)
			return new BigDecimal(value.doubleValue());
		if (type == BigInteger.class)
			return new BigInteger(value.toString());
		throw new ExcelParserException("Tipo numérico desconhecido.");
	}

	private static Object convertStringToNumber(String value, Class<?> type) throws ExcelParserException {
		if (type == double.class || type == Double.class)
			return Double.parseDouble(value);
		if (type == int.class || type == Integer.class)
			return Integer.parseInt(value);
		if (type == long.class || type == Long.class)
			return Long.parseLong(value);
		if (type == float.class || type == Float.class)
			return Float.parseFloat(value);
		if (type == short.class || type == Short.class)
			return Short.parseShort(value);
		if (type == byte.class || type == Byte.class)
			return Byte.parseByte(value);
		if (type == BigDecimal.class)
			return new BigDecimal(value);
		if (type == BigInteger.class)
			return new BigInteger(value);
		throw new ExcelParserException("Tipo numérico desconhecido.");
	}


}
