package com.flexibleexcel.util;

import com.flexibleexcel.exception.ExcelExportException;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * 反射工具类
 */
public class ReflectionUtil {

    /**
     * 获取字段值
     * @param obj 对象实例
     * @param field 字段
     * @return 字段值
     */
    public static Object getFieldValue(Object obj, Field field) {
        try {
            field.setAccessible(true);
            return field.get(obj);
        } catch (IllegalAccessException e) {
            throw new ExcelExportException("无法访问字段: " + field.getName(), e);
        }
    }

    /**
     * 通过getter方法获取字段值
     * @param obj 对象实例
     * @param fieldName 字段名
     * @return 字段值
     */
    public static Object getValueByGetter(Object obj, String fieldName) {
        try {
            String getterName = "get" + capitalize(fieldName);
            Method getter = obj.getClass().getMethod(getterName);
            return getter.invoke(obj);
        } catch (Exception e) {
            // 尝试直接访问字段
            try {
                Field field = findField(obj.getClass(), fieldName);
                if (field != null) {
                    field.setAccessible(true);
                    return field.get(obj);
                }
            } catch (Exception ex) {
                // ignore
            }
            return null;
        }
    }

    /**
     * 首字母大写
     * @param str 输入字符串
     * @return 首字母大写后的字符串
     */
    public static String capitalize(String str) {
        if (str == null || str.isEmpty()) {
            return str;
        }
        return Character.toUpperCase(str.charAt(0)) + str.substring(1);
    }

    /**
     * 查找字段（包括父类）
     * @param clazz 类
     * @param fieldName 字段名
     * @return 字段，找不到返回null
     */
    public static Field findField(Class<?> clazz, String fieldName) {
        Class<?> current = clazz;
        while (current != null && current != Object.class) {
            try {
                return current.getDeclaredField(fieldName);
            } catch (NoSuchFieldException e) {
                current = current.getSuperclass();
            }
        }
        return null;
    }

    /**
     * 获取类所有字段（包括父类）
     * @param clazz 类
     * @return 所有字段列表
     */
    public static List<Field> getAllFields(Class<?> clazz) {
        List<Field> fields = new ArrayList<>();
        Class<?> current = clazz;
        while (current != null && current != Object.class) {
            fields.addAll(Arrays.asList(current.getDeclaredFields()));
            current = current.getSuperclass();
        }
        return fields;
    }

    /**
     * 格式化日期
     * @param date 日期对象
     * @param pattern 日期格式
     * @return 格式化后的字符串
     */
    public static String formatDate(Date date, String pattern) {
        if (date == null) {
            return "";
        }
        try {
            SimpleDateFormat sdf = new SimpleDateFormat(pattern);
            return sdf.format(date);
        } catch (Exception e) {
            return date.toString();
        }
    }

    /**
     * 格式化数字
     * @param value 数字值
     * @param format 格式
     * @return 格式化后的字符串
     */
    public static String formatNumber(Object value, String format) {
        if (value == null) {
            return "";
        }
        if (value instanceof Number) {
            try {
                if (format != null && !format.isEmpty()) {
                    return new java.text.DecimalFormat(format).format(value);
                }
                return value.toString();
            } catch (Exception e) {
                return value.toString();
            }
        }
        return value.toString();
    }

    /**
     * 判断是否为简单类型
     * @param type 要判断的类型
     * @return 是否为简单类型
     */
    public static boolean isSimpleType(Class<?> type) {
        return type.isPrimitive() ||
                type == String.class ||
                type == Integer.class ||
                type == Long.class ||
                type == Double.class ||
                type == Float.class ||
                type == Boolean.class ||
                type == Date.class ||
                type == java.sql.Date.class ||
                type == LocalDate.class ||
                type == LocalDateTime.class ||
                Number.class.isAssignableFrom(type) ||
                CharSequence.class.isAssignableFrom(type);
    }

    /**
     * 将对象转换为字符串
     * @param value 要转换的对象
     * @param dateFormat 日期格式
     * @return 字符串表示
     */
    public static String toString(Object value, String dateFormat) {
        if (value == null) {
            return "";
        }
        if (value instanceof Date) {
            return formatDate((Date) value, dateFormat);
        }
        if (value instanceof LocalDate) {
            return formatLocalDate((LocalDate) value, dateFormat);
        }
        if (value instanceof LocalDateTime) {
            return formatLocalDateTime((LocalDateTime) value, dateFormat);
        }
        return value.toString();
    }

    /**
     * 格式化LocalDate
     * @param date LocalDate对象
     * @param pattern 日期格式
     * @return 格式化后的字符串
     */
    public static String formatLocalDate(LocalDate date, String pattern) {
        if (date == null) {
            return "";
        }
        try {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(pattern);
            return date.format(formatter);
        } catch (Exception e) {
            return date.toString();
        }
    }

    /**
     * 格式化LocalDateTime
     * @param dateTime LocalDateTime对象
     * @param pattern 日期格式
     * @return 格式化后的字符串
     */
    public static String formatLocalDateTime(LocalDateTime dateTime, String pattern) {
        if (dateTime == null) {
            return "";
        }
        try {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(pattern);
            return dateTime.format(formatter);
        } catch (Exception e) {
            return dateTime.toString();
        }
    }

    /**
     * 将Map的entry转换为键值对列表
     * @param map 输入的Map
     * @return 键值对列表
     */
    public static List<Map.Entry<Object, Object>> mapToEntryList(Map<?, ?> map) {
        if (map == null) {
            return Collections.emptyList();
        }
        List<Map.Entry<Object, Object>> entries = new ArrayList<>();
        for (Map.Entry<?, ?> entry : map.entrySet()) {
            entries.add(new AbstractMap.SimpleEntry<>(entry.getKey(), entry.getValue()));
        }
        return entries;
    }
}
