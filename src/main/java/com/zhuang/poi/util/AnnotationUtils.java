package com.zhuang.poi.util;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;

import com.zhuang.poi.excel.ExcelColumn;


public class AnnotationUtils {

	public static Method getSetMethodByColumnName(Class<?> clazz, String columnName)
			throws NoSuchMethodException, SecurityException, IntrospectionException {
		Method result = null;
		for (Field field : clazz.getDeclaredFields()) {
			field.setAccessible(true);
			ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
			if (excelColumn == null)
				continue;
			if (excelColumn.name().equals(columnName)) {
				PropertyDescriptor propertyDescriptor = new PropertyDescriptor(field.getName(), clazz);
				result = propertyDescriptor.getWriteMethod();
				break;
			}
		}
		return result;
	}

}
