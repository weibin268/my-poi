package com.zhuang.poi.util;

public class ConvertUtils {

    public static Object changeType(String value, Class<?> clazz) {
        Object result;
        if (clazz == Short.class) {
            result = Short.parseShort(value);
        } else if (clazz == Integer.class) {
            result = Integer.parseInt(value);
        } else if (clazz == Long.class) {
            result = Long.parseLong(value);
        } else if (clazz == Float.class) {
            result = Float.parseFloat(value);
        } else if (clazz == Double.class) {
            result = Double.parseDouble(value);
        } else if (clazz == Boolean.class) {
            result = Boolean.parseBoolean(value);
        } else {
            result = value;
        }
        return result;
    }

}
