package com.flexibleexcel.core;

import com.flexibleexcel.annotation.Excel;
import com.flexibleexcel.annotation.ExcelConfig;
import com.flexibleexcel.core.ExcelContext.ExportField;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;
import java.util.Map;

/**
 * 导出协作支持对象。
 * 负责封装值解析、样式应用和单元格访问等通用能力。
 */
public class ExportSupport {
    private final StyleHelper styleHelper = new StyleHelper();
    private final ValueResolver valueResolver = new ValueResolver();

    public ExportSupport() {
    }

    /**
     * 计算当前数据行需要占用的最大行数。
     */
    public int getMaxListSize(Object data, ExportState state) {
        int maxSize = 1;
        for (ExportField field : state.context.getExportFields()) {
            if (field.isList()) {
                List<?> list = getListValue(data, field);
                if (list != null) {
                    maxSize = Math.max(maxSize, list.size());
                }
            }
        }
        return maxSize;
    }

    /**
     * 计算 Map 字段展开后的最大列数。
     */
    public int getMaxMapColumns(List<?> dataList, ExportField field) {
        int maxColumns = 1;
        for (Object data : dataList) {
            Map<?, ?> map = getMapValue(data, field);
            if (map != null) {
                maxColumns = Math.max(maxColumns, map.size());
            }
        }
        return maxColumns;
    }

    /**
     * 获取 Map 字段值。
     */
    public Map<?, ?> getMapValue(Object data, ExportField field) {
        return valueResolver.getMapValue(data, field);
    }

    /**
     * 获取 List 字段值。
     */
    public List<?> getListValue(Object data, ExportField field) {
        return valueResolver.getListValue(data, field);
    }

    /**
     * 获取字段表头文本。
     */
    public String getHeaderText(ExportField field) {
        if (field.getExcel() != null && !field.getExcel().header().isEmpty()) {
            return field.getExcel().header();
        }
        return field.getField().getName();
    }

    /**
     * 获取普通字段导出值。
     */
    public String getFieldValue(Object data, ExportField field) {
        return valueResolver.getFieldValue(data, field);
    }

    /**
     * 获取 List 元素的导出值。
     */
    public String getListItemValue(Object item) {
        return valueResolver.getListItemValue(item);
    }

    /**
     * 获取 List 元素子字段的导出值。
     */
    public String getListItemFieldValue(Object item, ExportField field) {
        return valueResolver.getListItemFieldValue(item, field);
    }

    /**
     * 确保目标单元格存在并返回。
     */
    public XSSFCell ensureRowAndGetCell(XSSFSheet sheet, int rowIndex, int columnIndex) {
        XSSFRow row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        XSSFCell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    /**
     * 为合并区域应用边框。
     */
    public void applyBordersToMergedRegion(XSSFSheet sheet, CellRangeAddress region) {
        styleHelper.applyBordersToMergedRegion(sheet, region);
    }

    /**
     * 为表头合并区域应用边框。
     */
    public void applyHeaderBordersToMergedRegion(XSSFSheet sheet, CellRangeAddress region) {
        styleHelper.applyHeaderBordersToMergedRegion(sheet, region);
    }

    /**
     * 应用表头样式。
     */
    public void applyHeaderStyle(XSSFCell cell) {
        styleHelper.applyHeaderStyle(cell);
    }

    /**
     * 应用嵌套主表头样式。
     */
    public void applyNestedMainHeaderStyle(XSSFCell cell, ExcelConfig config, Excel excel) {
        styleHelper.applyNestedMainHeaderStyle(cell, config, excel);
    }

    /**
     * 应用嵌套次级表头样式。
     */
    public void applyNestedSubHeaderStyle(XSSFCell cell) {
        styleHelper.applyNestedSubHeaderStyle(cell);
    }

    /**
     * 对整行应用表头样式。
     */
    public void applyHeaderStyleToRow(XSSFRow row, ExcelConfig config) {
        styleHelper.applyHeaderStyleToRow(row, config);
    }

    /**
     * 对整行应用次级表头样式。
     */
    public void applySubHeaderStyleToRow(XSSFRow row) {
        styleHelper.applySubHeaderStyleToRow(row);
    }

    /**
     * 应用数据单元格样式。
     */
    public void applyCellStyle(XSSFCell cell, Excel excel) {
        styleHelper.applyCellStyle(cell, excel);
    }
}
