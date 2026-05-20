package com.flexibleexcel.processor;

import com.flexibleexcel.annotation.Excel;
import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.util.ReflectionUtil;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

/**
 * 嵌套对象处理器
 * 处理嵌套对象的横向展开
 */
public class NestedProcessor {

    /**
     * 嵌套字段信息（包含主字段和子字段）
     */
    public static class NestedFieldInfo {
        /**
         * 主字段（带@Excel注解的嵌套对象字段）
         */
        public final ExportField parentField;

        /**
         * 主表头文本
         */
        public final String parentHeader;

        /**
         * 主表头占据的起始列
         */
        public final int startColumn;

        /**
         * 子字段列表（嵌套对象的可导出字段）
         */
        public final List<ExportField> childFields;

        /**
         * 次级表头文本列表
         */
        public final List<String> childHeaders;

        /**
         * 该嵌套字段占据的列数
         */
        public final int columnSpan;

        public NestedFieldInfo(ExportField parentField, String parentHeader,
                               int startColumn, List<ExportField> childFields,
                               List<String> childHeaders) {
            this.parentField = parentField;
            this.parentHeader = parentHeader;
            this.startColumn = startColumn;
            this.childFields = childFields;
            this.childHeaders = childHeaders;
            this.columnSpan = childFields.size();
        }
    }

    /**
     * 从数据列表中收集所有嵌套对象展开后的次级表头
     *
     * @param dataList 数据列表
     * @param field    嵌套字段
     * @return 所有次级表头的key列表（去重）
     */
    public static List<String> getAllChildHeaders(List<?> dataList, ExportField field) {
        Set<String> headerSet = new LinkedHashSet<>();

        for (Object data : dataList) {
            try {
                Object nestedObj = ReflectionUtil.getFieldValue(data, field.getField());
                if (nestedObj == null) {
                    continue;
                }

                List<ExportField> childFields = getChildFields(nestedObj.getClass(), field.getExcel().nested());
                for (ExportField childField : childFields) {
                    String header = getChildHeader(childField);
                    headerSet.add(header);
                }
            } catch (Exception e) {
                // ignore
            }
        }

        return new ArrayList<>(headerSet);
    }

    /**
     * 获取嵌套对象的子字段
     *
     * @param clazz     嵌套对象的类
     * @param nestedMode 嵌套模式
     * @return 子字段列表
     */
    public static List<ExportField> getChildFields(Class<?> clazz, Excel.NestedMode nestedMode) {
        List<ExportField> childFields = new ArrayList<>();
        List<Field> allFields = ReflectionUtil.getAllFields(clazz);

        int columnIndex = 0;
        for (Field f : allFields) {
            // 跳过@ExcelIgnore标注的字段
            if (f.isAnnotationPresent(com.flexibleexcel.annotation.ExcelIgnore.class)) {
                continue;
            }

            Excel excel = f.getAnnotation(Excel.class);
            if (excel != null) {
                // 递归模式：子字段也是嵌套对象，需要进一步展开
                Class<?> fieldType = f.getType();
                boolean isNested = !isSimpleType(fieldType) &&
                                   !List.class.isAssignableFrom(fieldType) &&
                                   !Map.class.isAssignableFrom(fieldType) &&
                                   excel.nested() != Excel.NestedMode.NONE;

                childFields.add(new ExportField(f, excel, columnIndex++, isNested, excel.nested()));
            }
        }

        return childFields;
    }

    /**
     * 获取子字段的表头文本
     */
    private static String getChildHeader(ExportField field) {
        if (field.getExcel() != null && !field.getExcel().header().isEmpty()) {
            return field.getExcel().header();
        }
        return field.getField().getName();
    }

    /**
     * 检查是否为简单类型
     */
    private static boolean isSimpleType(Class<?> type) {
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
     * 检查类型是否为嵌套对象
     */
    public static boolean isNestedObject(Class<?> type) {
        if (isSimpleType(type)) {
            return false;
        }
        if (List.class.isAssignableFrom(type) || Map.class.isAssignableFrom(type)) {
            return false;
        }
        // 检查是否有@Excel标注的字段
        List<Field> fields = ReflectionUtil.getAllFields(type);
        for (Field f : fields) {
            if (f.isAnnotationPresent(Excel.class) &&
                !f.isAnnotationPresent(com.flexibleexcel.annotation.ExcelIgnore.class)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 构建嵌套字段信息
     *
     * @param field       主字段
     * @param startColumn 起始列
     * @param dataList    数据列表
     * @return 嵌套字段信息
     */
    public static NestedFieldInfo buildNestedFieldInfo(ExportField field, int startColumn, List<?> dataList) {
        String parentHeader = "";
        if (field.getExcel() != null && !field.getExcel().header().isEmpty()) {
            parentHeader = field.getExcel().header();
        } else {
            parentHeader = field.getField().getName();
        }

        // 获取子字段
        List<ExportField> childFields = getChildFields(field.getField().getType(), field.getExcel().nested());

        // 收集所有次级表头
        List<String> childHeaders = new ArrayList<>();
        if (field.getExcel().nested() == Excel.NestedMode.HORIZONTAL) {
            // HORIZONTAL模式：收集所有数据中的次级表头
            childHeaders = getAllChildHeaders(dataList, field);
        } else {
            // RECURSIVE模式：直接使用子字段的表头
            for (ExportField childField : childFields) {
                childHeaders.add(getChildHeader(childField));
            }
        }

        return new NestedFieldInfo(field, parentHeader, startColumn, childFields, childHeaders);
    }

    /**
     * 获取嵌套对象的值
     *
     * @param data        数据对象
     * @param field       嵌套字段
     * @param childHeader 次级表头（用于HORIZONTAL模式匹配）
     * @return 值
     */
    public static Object getNestedValue(Object data, ExportField field, String childHeader) {
        try {
            Object nestedObj = ReflectionUtil.getFieldValue(data, field.getField());
            if (nestedObj == null) {
                return null;
            }

            if (field.getExcel().nested() == Excel.NestedMode.RECURSIVE) {
                // RECURSIVE模式：直接在嵌套对象中查找对应次级表头的字段
                return getRecursiveValue(nestedObj, childHeader);
            } else {
                // HORIZONTAL模式：通过次级表头匹配子字段
                return getHorizontalValue(nestedObj, childHeader);
            }
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * RECURSIVE模式：根据次级表头获取值
     */
    private static Object getRecursiveValue(Object nestedObj, String childHeader) {
        List<ExportField> childFields = getChildFields(nestedObj.getClass(), Excel.NestedMode.RECURSIVE);
        for (ExportField childField : childFields) {
            String header = getChildHeader(childField);
            if (header.equals(childHeader)) {
                return ReflectionUtil.getFieldValue(nestedObj, childField.getField());
            }
        }
        return null;
    }

    /**
     * HORIZONTAL模式：在嵌套对象中查找匹配次级表头的字段
     */
    private static Object getHorizontalValue(Object nestedObj, String childHeader) {
        List<ExportField> childFields = getChildFields(nestedObj.getClass(), Excel.NestedMode.HORIZONTAL);
        for (ExportField childField : childFields) {
            String header = getChildHeader(childField);
            if (header.equals(childHeader)) {
                return ReflectionUtil.getFieldValue(nestedObj, childField.getField());
            }
        }
        return null;
    }
}
