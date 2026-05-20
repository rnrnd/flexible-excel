package com.flexibleexcel.util;

import com.flexibleexcel.exception.ExcelExportException;

import java.lang.reflect.Field;
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

}
