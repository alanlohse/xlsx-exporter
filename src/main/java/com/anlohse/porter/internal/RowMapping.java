package com.anlohse.porter.internal;

import java.io.Serializable;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.ConcurrentHashMap;

import com.anlohse.porter.ExcelParserException;
import com.anlohse.porter.annotation.Column;

public class RowMapping<E> implements Serializable {

	private static final long serialVersionUID = 1L;

	static ConcurrentHashMap<Class<?>, RowMapping<?>> mappings = new ConcurrentHashMap<>();

	public static <T> RowMapping<T> getMapping(Class<T> clazz) {
		@SuppressWarnings("unchecked")
		RowMapping<T> mapping = (RowMapping<T>) mappings.get(clazz);
		if (mapping == null) {
			mapping = RowMapping.create(clazz);
			mappings.put(clazz, mapping);
		}
		return mapping;
	}

	public static <T> RowMapping<T> create(Class<T> clazz) {
		List<ColMapping> cols = new ArrayList<>();
		List<Field> fields = getFields(clazz);
		int maxCol = 0;
		for (int i = 0; i < fields.size(); i++) {
			Field field = fields.get(i);
			Column col = field.getAnnotation(Column.class);
			if (col != null) {
				Method getter = getGetter(clazz, field.getName());
				Method setter = getSetter(clazz, field.getName(), field.getType());
				ColMapping map = new ColMapping(field.getType(), col.header().isEmpty() ? field.getName() : col.header(), field.getName(), getter, setter, field, col.index(), i);
				maxCol = max(col.index(), i, maxCol);
				cols.add(map);
			}
		}
		Collections.sort(cols);
		return new RowMapping<T>(cols, clazz, maxCol);
	}

	private static int max(int a, int b, int c) {
		return Math.max(a, Math.max(b, c));
	}

	private static Method getSetter(Class<?> clazz, String name, Class<?> type) {
		try {
			return clazz.getMethod("set" + name.substring(0, 1).toUpperCase() + name.substring(1), type);
		} catch (Exception e) {
		}
		return null;
	}

	private static Method getGetter(Class<?> clazz, String name) {
		try {
			return clazz.getMethod(
					(clazz == boolean.class ? "is" : "get") + name.substring(0, 1).toUpperCase() + name.substring(1));
		} catch (Exception e) {
		}
		return null;
	}

	private static List<Field> getFields(Class<?> clazz) {
		List<Field> result = new ArrayList<>();
		if (clazz != Object.class) {
			Class<?> superClass = clazz.getSuperclass();
			result.addAll(getFields(superClass));
		}
		result.addAll(Arrays.asList(clazz.getDeclaredFields()));
		return result;
	}

	final List<ColMapping> cols;
	final Class<E> clazz;
	final int maxCol;

	private RowMapping(List<ColMapping> cols, Class<E> clazz, int maxCol) {
		this.cols = cols;
		this.clazz = clazz;
		this.maxCol = maxCol;
	}

	public int getCount() {
		return cols.size();
	}

	public E createRow(Object[] values, Locale locale) throws ExcelParserException {
		try {
			return doCreate(values, locale);
		} catch (Exception e) {
			throw new ExcelParserException(e);
		}
	}

	private E doCreate(Object[] values, Locale locale) throws Exception {
		E obj = clazz.newInstance();
		int i = 0;
		for (ColMapping col : cols) {
			if (col.index == i)
				i++;
			Object value = ConvertOps.convert(
					col.index != -1 ? (col.index < values.length ? values[col.index] : null) : values[i++],
					col.type, locale);
			if (col.setter == null) {
				col.field.setAccessible(true);
				col.field.set(obj, value);
			} else
				col.setter.invoke(obj, value);
		}
		return obj;
	}

	public Object[] getValues(E obj) throws Exception {
		Object[] values = new Object[cols.size()];
		int i = 0;
		for (ColMapping col : cols) {
			Object value;
			if (col.index == i)
				i++;
			if (col.getter == null) {
				col.field.setAccessible(true);
				value = col.field.get(obj);
			} else
				value = col.getter.invoke(obj);
			values[col.index != -1 ? col.index : i++] = value;
		}
		return values;
	}

	private static class ColMapping implements Serializable, Comparable<ColMapping> {

		private static final long serialVersionUID = 1L;

		private final Class<?> type;
		private final String header;
		private final Method getter;
		private final Method setter;
		private final Field field;
		private final int index;
		private final int forder;
		private final String colName;

		public ColMapping(Class<?> type, String header, String colName, Method getter, Method setter, Field field, int index,
				int forder) {
			this.type = type;
			this.header = header;
			this.getter = getter;
			this.setter = setter;
			this.field = field;
			this.index = index;
			this.forder = forder;
			this.colName = colName;
		}

		@Override
		public int compareTo(ColMapping o) {
			if (index != -1 && o.index != -1)
				return index - o.index;
			if (index == -1 && o.index == -1)
				return forder - o.forder;
			if (index != -1 && o.index == -1)
				return index - o.forder;
			if (index == -1 && o.index != -1)
				return forder - o.index;
			return 0;
		}

	}

	public int getIndex(int c) {
		return cols.get(c).index;
	}

	public String getHeader(int c) {
		return cols.get(c).header;
	}

	public String getColumnName(int c) {
		return cols.get(c).colName;
	}

}
