package com.flexibleexcel.processor;

import com.flexibleexcel.annotation.Excel;
import com.flexibleexcel.core.CellData;
import com.flexibleexcel.core.ExcelContext;
import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.util.ReflectionUtil;

import java.util.ArrayList;
import java.util.List;

/**
 * 字段处理器
 * 处理普通字段（非List、非Map）
 */
public class FieldProcessor {

    private final ExcelContext context;

    public FieldProcessor(ExcelContext context) {
        this.context = context;
    }

    /**
     * 处理普通字段
     *
     * @param data      数据对象
     * @param field     字段信息
     * @param row       行索引
     * @param column    列索引
     * @param rowSpan   该字段需要跨越的行数（用于非List字段的纵向合并）
     * @return 单元格数据
     */
    public CellData process(Object data, ExportField field, int row, int column, int rowSpan) {
        try {
            Object value = ReflectionUtil.getFieldValue(data, field.getField());
            String text = formatValue(value, field);
            return CellData.merged(text, row, row + rowSpan - 1, column, column);
        } catch (Exception e) {
            return CellData.merged("Error: " + e.getMessage(), row, row + rowSpan - 1, column, column);
        }
    }

    /**
     * 格式化字段值
     */
    private String formatValue(Object value, ExportField field) {
        if (value == null) {
            return "";
        }

        Excel excel = field.getExcel();
        if (excel == null) {
            return value.toString();
        }

        // 日期格式化
        if (value instanceof java.util.Date) {
            return ReflectionUtil.formatDate((java.util.Date) value, excel.dateFormat());
        }

        // 数字格式化
        if (value instanceof Number && !excel.numberFormat().isEmpty()) {
            return ReflectionUtil.formatNumber(value, excel.numberFormat());
        }

        return value.toString();
    }

    /**
     * 创建表头单元格数据
     *
     * @param field        字段信息
     * @param row          行索引
     * @param column       列索引
     * @param columnSpan   列跨越数（用于Map类型的表头合并）
     * @return 表头单元格数据
     */
    public CellData createHeader(ExportField field, int row, int column, int columnSpan) {
        String headerText = "";
        if (field.getExcel() != null) {
            headerText = field.getExcel().header();
        }
        if (headerText.isEmpty()) {
            headerText = field.getField().getName();
        }

        if (columnSpan > 1) {
            return CellData.mergedHeader(headerText, row, row, column, column + columnSpan - 1);
        } else {
            return CellData.header(headerText, row, column);
        }
    }

    /**
     * 获取需要合并的行数
     * 对于List字段，返回1；对于非List字段，返回最大List长度
     */
    public int getRowSpan(ExportField field, List<?> dataList) {
        if (field.isList()) {
            return 1; // List字段每个元素占一行
        }

        // 对于非List字段，计算最大List长度
        int maxListSize = 1;
        for (Object data : dataList) {
            for (ExportField ef : context.getExportFields()) {
                if (ef.isList()) {
                    int listSize = getListSize(data, ef);
                    maxListSize = Math.max(maxListSize, listSize);
                }
            }
        }
        return maxListSize;
    }

    /**
     * 获取List字段的长度
     */
    private int getListSize(Object data, ExportField field) {
        try {
            Object value = ReflectionUtil.getFieldValue(data, field.getField());
            if (value instanceof List) {
                return ((List<?>) value).size();
            }
        } catch (Exception e) {
            // ignore
        }
        return 1;
    }
}
