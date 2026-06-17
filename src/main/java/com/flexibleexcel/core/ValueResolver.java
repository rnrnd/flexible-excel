package com.flexibleexcel.core;

import com.flexibleexcel.annotation.Excel;
import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.util.ReflectionUtil;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * 值解析器，负责反射取值与格式化。
 */
class ValueResolver {

    /**
     * 获取字段值。
     */
    String getFieldValue(Object data, ExportField field) {
        try {
            Object value = ReflectionUtil.getFieldValue(data, field.getField());
            if (value == null) {
                return "";
            }

            Excel excel = field.getExcel();
            if (excel != null) {
                if (value instanceof Date) {
                    return ReflectionUtil.formatDate((Date) value, excel.dateFormat());
                }
                if (value instanceof LocalDate) {
                    return ReflectionUtil.formatLocalDate((LocalDate) value, excel.dateFormat());
                }
                if (value instanceof LocalDateTime) {
                    return ReflectionUtil.formatLocalDateTime((LocalDateTime) value, excel.dateFormat());
                }
                if (value instanceof Number && !excel.numberFormat().isEmpty()) {
                    return ReflectionUtil.formatNumber(value, excel.numberFormat());
                }
            }

            return value.toString();
        } catch (Exception e) {
            return "Error: " + e.getMessage();
        }
    }

    /**
     * 获取 List 元素的展示值。
     */
    String getListItemValue(Object item) {
        if (item == null) {
            return "";
        }

        Class<?> itemType = item.getClass();
        java.lang.reflect.Field[] fields = itemType.getDeclaredFields();

        List<String> values = new ArrayList<>();
        for (java.lang.reflect.Field f : fields) {
            Excel excel = f.getAnnotation(Excel.class);
            if (excel != null) {
                try {
                    f.setAccessible(true);
                    Object value = f.get(item);
                    if (value != null) {
                        if (value instanceof Date) {
                            values.add(ReflectionUtil.formatDate((Date) value, excel.dateFormat()));
                        } else if (value instanceof Number && !excel.numberFormat().isEmpty()) {
                            values.add(ReflectionUtil.formatNumber(value, excel.numberFormat()));
                        } else {
                            values.add(value.toString());
                        }
                    } else {
                        values.add("");
                    }
                } catch (Exception e) {
                    values.add("");
                }
            }
        }

        if (values.isEmpty()) {
            return item.toString();
        }
        return String.join(", ", values);
    }

    /**
     * 获取 List 元素指定字段的值。
     */
    String getListItemFieldValue(Object item, ExportField field) {
        if (item == null) {
            return "";
        }
        try {
            java.lang.reflect.Field f = field.getField();
            f.setAccessible(true);
            Object value = f.get(item);
            if (value == null) {
                return "";
            }
            Excel excel = field.getExcel();
            if (excel != null) {
                if (value instanceof Date) {
                    return ReflectionUtil.formatDate((Date) value, excel.dateFormat());
                }
                if (value instanceof Number && !excel.numberFormat().isEmpty()) {
                    return ReflectionUtil.formatNumber(value, excel.numberFormat());
                }
            }
            return value.toString();
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * 获取 Map 值。
     */
    Map<?, ?> getMapValue(Object data, ExportField field) {
        try {
            Object value = ReflectionUtil.getFieldValue(data, field.getField());
            if (value instanceof Map) {
                return (Map<?, ?>) value;
            }
        } catch (Exception e) {
            // ignore
        }
        return null;
    }

    /**
     * 获取 List 值。
     */
    List<?> getListValue(Object data, ExportField field) {
        try {
            Object value = ReflectionUtil.getFieldValue(data, field.getField());
            if (value instanceof List) {
                return (List<?>) value;
            }
        } catch (Exception e) {
            // ignore
        }
        return null;
    }
}
