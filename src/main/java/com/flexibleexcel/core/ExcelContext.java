package com.flexibleexcel.core;

import com.flexibleexcel.annotation.Excel;
import lombok.Data;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;

/**
 * Excel导出上下文
 * 存储导出过程中的配置信息和中间数据
 */
@Data
public class ExcelContext {

    /**
     * 数据列表
     */
    private List<?> dataList;

    /**
     * 导出的字段列表
     */
    private List<ExportField> exportFields;

    /**
     * 最大行数
     */
    private int maxRows;

    /**
     * 最大列数
     */
    private int maxColumns;

    /**
     * 表头行数
     */
    private int headerRows;

    /**
     * 数据起始行
     */
    private int dataStartRow;

    /**
     * 数据起始列
     */
    private int dataStartColumn;

    /**
     * 列宽映射
     */
    private int[] columnWidths;

    /**
     * 行高映射
     */
    private int[] rowHeights;

    /**
     * 导出字段信息
     */
    @Data
    public static class ExportField {
        private Field field;
        private Excel excel;
        private int columnIndex;
        private Class<?> fieldType;
        private boolean isList;
        private boolean isMap;
        private boolean isNested;
        private Excel.NestedMode nestedMode;

        public ExportField(Field field, Excel excel, int columnIndex) {
            this.field = field;
            this.excel = excel;
            this.columnIndex = columnIndex;
            this.fieldType = field.getType();
            this.isList = List.class.isAssignableFrom(fieldType);
            this.isMap = java.util.Map.class.isAssignableFrom(fieldType);
            this.isNested = isNestedObject(field.getType(), excel);
            this.nestedMode = (excel != null) ? excel.nested() : Excel.NestedMode.NONE;
        }

        public ExportField(Field field, Excel excel, int columnIndex, boolean isNested, Excel.NestedMode nestedMode) {
            this.field = field;
            this.excel = excel;
            this.columnIndex = columnIndex;
            this.fieldType = field.getType();
            this.isList = List.class.isAssignableFrom(fieldType);
            this.isMap = java.util.Map.class.isAssignableFrom(fieldType);
            this.isNested = isNested;
            this.nestedMode = nestedMode;
        }

        public int getOrder() {
            return excel != null ? excel.order() : 0;
        }

        /**
         * 判断是否为嵌套对象
         */
        private static boolean isNestedObject(Class<?> type, Excel excel) {
            if (type == null || excel == null) {
                return false;
            }
            if (isSimpleType(type)) {
                return false;
            }
            if (List.class.isAssignableFrom(type) || java.util.Map.class.isAssignableFrom(type)) {
                return false;
            }
            return excel.nested() != Excel.NestedMode.NONE;
        }

        /**
         * 判断是否为简单类型
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
                    Number.class.isAssignableFrom(type) ||
                    CharSequence.class.isAssignableFrom(type);
        }
    }

    /**
     * 创建上下文
     */
    public static ExcelContext create(Class<?> clazz, List<?> dataList) {
        ExcelContext context = new ExcelContext();
        context.setDataList(dataList);
        context.setExportFields(extractExportFields(clazz));
        context.setHeaderRows(1);
        context.setDataStartRow(1);
        context.setDataStartColumn(0);
        return context;
    }

    /**
     * 提取可导出的字段
     */
    private static List<ExportField> extractExportFields(Class<?> clazz) {
        List<ExportField> fields = new ArrayList<>();
        int columnIndex = 0;

        // 获取所有带@Excel注解的字段
        for (Field field : clazz.getDeclaredFields()) {
            if (field.isAnnotationPresent(Excel.class)) {
                Excel excel = field.getAnnotation(Excel.class);
                fields.add(new ExportField(field, excel, columnIndex));
                columnIndex++;
            }
        }

        // 按order排序
        fields.sort(Comparator.comparingInt(ExportField::getOrder));

        return fields;
    }

    /**
     * 初始化列宽数组
     */
    public void initColumnWidths() {
        int size = exportFields.size();
        if (size == 0) {
            this.columnWidths = new int[0];
            return;
        }

        this.columnWidths = new int[size];
        for (int i = 0; i < size; i++) {
            ExportField field = exportFields.get(i);
            int width = field.getExcel() != null ? field.getExcel().columnWidth() : -1;
            this.columnWidths[i] = width;
        }
    }

    /**
     * 初始化行高数组
     */
    public void initRowHeights(int rows) {
        this.rowHeights = new int[rows];
    }
}
