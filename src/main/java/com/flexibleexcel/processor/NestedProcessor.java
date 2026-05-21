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
 * 处理嵌套对象的横向展开（支持任意嵌套深度）
 */
public class NestedProcessor {

    /**
     * 嵌套字段信息（支持递归嵌套结构）
     */
    public static class NestedFieldInfo {
        /**
         * 主字段（带@Excel注解的嵌套对象字段）
         */
        public final ExportField parentField;

        /**
         * 主表头文本（当前深度的表头）
         */
        public final String parentHeader;

        /**
         * 主表头占据的起始列
         */
        public final int startColumn;

        /**
         * 该嵌套字段占用的叶子列总数
         */
        public final int columnSpan;

        /**
         * 直接子列定义（出现在下一深度的列）
         */
        public final List<NestedColumn> columns;

        /**
         * 直接子字段列表（向后兼容，扁平视图）
         */
        public final List<ExportField> childFields;

        /**
         * 直接子表头文本列表（向后兼容，扁平视图）
         */
        public final List<String> childHeaders;

        public NestedFieldInfo(ExportField parentField, String parentHeader,
                               int startColumn, int columnSpan,
                               List<NestedColumn> columns,
                               List<ExportField> childFields,
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
            /**
             * 该列对应的字段
             */
            public final ExportField field;

            /**
             * 该列在当前深度的表头文本
             */
            public final String header;

            /**
             * 该列占用的叶子列数（简单字段为1，嵌套字段为其展开后的列数）
             */
            public final int leafColumnSpan;

            /**
             * 如果该列本身也是嵌套对象，则包含其展开信息；否则为null
             */
            public final NestedFieldInfo nestedInfo;

            public NestedColumn(ExportField field, String header, int leafColumnSpan, NestedFieldInfo nestedInfo) {
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
     *
     * @param clazz     嵌套对象的类
     * @return 子字段列表
     */
    public static List<ExportField> getChildFields(Class<?> clazz) {
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
     * 构建嵌套字段信息（递归支持任意嵌套深度）
     *
     * @param field       主字段
     * @param startColumn 起始列
     * @return 嵌套字段信息（包含递归子结构）
     */
    public static NestedFieldInfo buildNestedFieldInfo(ExportField field, int startColumn) {
        String parentHeader;
        if (field.getExcel() != null && !field.getExcel().header().isEmpty()) {
            parentHeader = field.getExcel().header();
        } else {
            parentHeader = field.getField().getName();
        }

        List<ExportField> childFields = getChildFields(field.getField().getType());

        List<NestedFieldInfo.NestedColumn> columns = new ArrayList<>();
        List<String> childHeaders = new ArrayList<>();
        int totalColumnSpan = 0;

        for (ExportField childField : childFields) {
            String childHeader = getChildHeader(childField);
            childHeaders.add(childHeader);

            if (childField.isNested() && childField.getExcel().nested() == Excel.NestedMode.HORIZONTAL) {
                // 子字段也是嵌套对象，递归构建
                NestedFieldInfo childNested = buildNestedFieldInfo(childField, startColumn + totalColumnSpan);
                columns.add(new NestedFieldInfo.NestedColumn(childField, childHeader, childNested.columnSpan, childNested));
                totalColumnSpan += childNested.columnSpan;
            } else {
                // 简单字段或不需要展开的字段
                columns.add(new NestedFieldInfo.NestedColumn(childField, childHeader, 1, null));
                totalColumnSpan += 1;
            }
        }

        return new NestedFieldInfo(field, parentHeader, startColumn, totalColumnSpan,
                                   columns, childFields, childHeaders);
    }

}
