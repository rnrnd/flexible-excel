package com.flexibleexcel.core.structure;

import com.flexibleexcel.annotation.Excel;
import com.flexibleexcel.util.ReflectionUtil;
import com.flexibleexcel.core.ExcelContext;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * 嵌套对象处理器
 * 处理嵌套对象的横向展开（支持任意嵌套深度）
 */
public class NestedProcessor {

    /**
     * 嵌套字段信息（支持递归嵌套结构）
     */
    public static class NestedFieldInfo {
        public final ExcelContext.ExportField parentField;
        public final String parentHeader;
        public final int startColumn;
        public final int columnSpan;
        public final List<NestedColumn> columns;
        public final List<ExcelContext.ExportField> childFields;
        public final List<String> childHeaders;

        public NestedFieldInfo(ExcelContext.ExportField parentField, String parentHeader,
                               int startColumn, int columnSpan,
                               List<NestedColumn> columns,
                               List<ExcelContext.ExportField> childFields,
                               List<String> childHeaders) {
            this.parentField = parentField;
            this.parentHeader = parentHeader;
            this.startColumn = startColumn;
            this.columnSpan = columnSpan;
            this.columns = columns;
            this.childFields = childFields;
            this.childHeaders = childHeaders;
        }

        /**
         * 嵌套列定义
         */
        public static class NestedColumn {
            public final ExcelContext.ExportField field;
            public final String header;
            public final int leafColumnSpan;
            public final NestedFieldInfo nestedInfo;

            public NestedColumn(ExcelContext.ExportField field, String header, int leafColumnSpan, NestedFieldInfo nestedInfo) {
                this.field = field;
                this.header = header;
                this.leafColumnSpan = leafColumnSpan;
                this.nestedInfo = nestedInfo;
            }

            public boolean isNested() {
                return nestedInfo != null;
            }
        }
    }

    /**
     * 获取嵌套对象的子字段
     */
    public static List<ExcelContext.ExportField> getChildFields(Class<?> clazz) {
        List<ExcelContext.ExportField> childFields = new ArrayList<>();
        List<Field> allFields = ReflectionUtil.getAllFields(clazz);

        int columnIndex = 0;
        for (Field f : allFields) {
            if (f.isAnnotationPresent(com.flexibleexcel.annotation.ExcelIgnore.class)) {
                continue;
            }

            Excel excel = f.getAnnotation(Excel.class);
            if (excel != null) {
                Class<?> fieldType = f.getType();
                boolean isNested = !isSimpleType(fieldType)
                        && !List.class.isAssignableFrom(fieldType)
                        && !Map.class.isAssignableFrom(fieldType)
                        && excel.nested() != Excel.NestedMode.NONE;

                childFields.add(new ExcelContext.ExportField(f, excel, columnIndex++, isNested, excel.nested()));
            }
        }

        return childFields;
    }

    /**
     * 获取子字段的表头文本
     */
    private static String getChildHeader(ExcelContext.ExportField field) {
        if (field.getExcel() != null && !field.getExcel().header().isEmpty()) {
            return field.getExcel().header();
        }
        return field.getField().getName();
    }

    /**
     * 检查是否为简单类型
     */
    private static boolean isSimpleType(Class<?> type) {
        return type.isPrimitive()
                || type == String.class
                || type == Integer.class
                || type == Long.class
                || type == Double.class
                || type == Float.class
                || type == Boolean.class
                || type == Date.class
                || type == java.sql.Date.class
                || type == LocalDate.class
                || type == LocalDateTime.class
                || Number.class.isAssignableFrom(type)
                || CharSequence.class.isAssignableFrom(type);
    }

    /**
     * 构建嵌套字段信息（递归支持任意嵌套深度）
     */
    public static NestedFieldInfo buildNestedFieldInfo(ExcelContext.ExportField field, int startColumn) {
        String parentHeader;
        if (field.getExcel() != null && !field.getExcel().header().isEmpty()) {
            parentHeader = field.getExcel().header();
        } else {
            parentHeader = field.getField().getName();
        }

        List<ExcelContext.ExportField> childFields = getChildFields(field.getField().getType());

        List<NestedFieldInfo.NestedColumn> columns = new ArrayList<>();
        List<String> childHeaders = new ArrayList<>();
        int totalColumnSpan = 0;

        for (ExcelContext.ExportField childField : childFields) {
            String childHeader = getChildHeader(childField);
            childHeaders.add(childHeader);

            if (childField.isNested() && childField.getExcel().nested() == Excel.NestedMode.HORIZONTAL) {
                NestedFieldInfo childNested = buildNestedFieldInfo(childField, startColumn + totalColumnSpan);
                columns.add(new NestedFieldInfo.NestedColumn(childField, childHeader, childNested.columnSpan, childNested));
                totalColumnSpan += childNested.columnSpan;
            } else {
                columns.add(new NestedFieldInfo.NestedColumn(childField, childHeader, 1, null));
                totalColumnSpan += 1;
            }
        }

        return new NestedFieldInfo(field, parentHeader, startColumn, totalColumnSpan,
                columns, childFields, childHeaders);
    }
}
