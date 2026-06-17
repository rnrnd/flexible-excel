package com.flexibleexcel.core.write;

import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.core.ExportState;
import com.flexibleexcel.core.ExportSupport;
import com.flexibleexcel.core.structure.ListFieldInfo;
import com.flexibleexcel.core.structure.MapKeyCollector;
import com.flexibleexcel.core.structure.NestedProcessor;
import com.flexibleexcel.util.ReflectionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;
import java.util.Map;

/**
 * 数据写入协作对象，负责数据区逐行逐列渲染。
 */
public class DataWriter {
    private final ExportSupport support;
    private final XSSFSheet sheet;
    private final List<?> dataList;
    private final ExportState state;

    public DataWriter(ExportSupport support, XSSFSheet sheet, List<?> dataList, ExportState state) {
        this.support = support;
        this.sheet = sheet;
        this.dataList = dataList;
        this.state = state;
    }

    /**
     * 执行数据写入。
     */
    public void write() {
        int currentRow = state.context.getDataStartRow();

        for (Object data : dataList) {
            int maxListSize = support.getMaxListSize(data, state);
            int currentColumn = 0;

            for (ExportField field : state.context.getExportFields()) {
                switch (FieldKind.from(field)) {
                    case MAP:
                        currentColumn = writeMapDataCells(data, field, currentRow, currentColumn);
                        break;
                    case NESTED:
                        currentColumn = writeNestedOrScalarData(data, field, currentRow, currentColumn);
                        break;
                    case LIST:
                        currentColumn = writeListDataCells(data, field, currentRow, currentColumn);
                        break;
                    default:
                        currentColumn = writeScalarDataCell(currentRow, currentColumn, data, field);
                        break;
                }
            }

            currentRow += Math.max(maxListSize, 1);
        }
    }

    /**
     * 写入 Map 字段对应的数据列。
     */
    private int writeMapDataCells(Object data, ExportField field, int rowIndex, int columnIndex) {
        Map<?, ?> map = support.getMapValue(data, field);
        if (map != null && !map.isEmpty()) {
            List<?> allKeys = MapKeyCollector.getAllKeys(dataList, field);
            for (Object key : allKeys) {
                Object value = map.get(key);
                writeCell(rowIndex, columnIndex, value != null ? value.toString() : "", field.getExcel());
                columnIndex++;
            }
            return columnIndex;
        }

        return writeBlankCells(rowIndex, columnIndex, support.getMaxMapColumns(dataList, field), field.getExcel());
    }

    /**
     * 写入嵌套字段，若未命中结构信息则退回普通字段写法。
     */
    private int writeNestedOrScalarData(Object data, ExportField field, int rowIndex, int columnIndex) {
        NestedProcessor.NestedFieldInfo nestedInfo = state.getNestedFieldInfo(field);
        if (nestedInfo != null) {
            return writeNestedDataRecursive(data, nestedInfo, rowIndex, columnIndex);
        }

        Object value = ReflectionUtil.getFieldValue(data, field.getField());
        writeCell(rowIndex, columnIndex, value != null ? value.toString() : "", field.getExcel());
        return columnIndex + 1;
    }

    /**
     * 写入 List 字段对应的数据列。
     */
    private int writeListDataCells(Object data, ExportField field, int rowIndex, int columnIndex) {
        List<?> list = support.getListValue(data, field);
        ListFieldInfo listInfo = state.getListFieldInfo(field);

        if (listInfo != null && listInfo.columnSpan > 0) {
            return writeExpandedListDataCells(list, listInfo, rowIndex, columnIndex);
        }

        if (list != null && !list.isEmpty()) {
            for (int i = 0; i < list.size(); i++) {
                Object item = list.get(i);
                String value = support.getListItemValue(item);
                writeCell(rowIndex + i, columnIndex, value, field.getExcel());
            }
        }
        return columnIndex + 1;
    }

    /**
     * 递归写入嵌套对象数据。
     */
    private int writeNestedDataRecursive(Object data,
                                         NestedProcessor.NestedFieldInfo nestedInfo,
                                         int currentRow, int currentColumn) {
        Object nestedObj = null;
        try {
            nestedObj = ReflectionUtil.getFieldValue(data, nestedInfo.parentField.getField());
        } catch (Exception e) {
            // ignore
        }

        int col = currentColumn;
        for (NestedProcessor.NestedFieldInfo.NestedColumn colDef : nestedInfo.columns) {
            if (colDef.isNested()) {
                if (nestedObj != null) {
                    col = writeNestedDataRecursive(nestedObj, colDef.nestedInfo, currentRow, col);
                } else {
                    col = writeBlankCells(currentRow, col, colDef.leafColumnSpan, nestedInfo.parentField.getExcel());
                }
            } else {
                String value = "";
                if (nestedObj != null) {
                    Object fieldValue = ReflectionUtil.getFieldValue(nestedObj, colDef.field.getField());
                    value = fieldValue != null ? fieldValue.toString() : "";
                }
                writeCell(currentRow, col, value, nestedInfo.parentField.getExcel());
                col++;
            }
        }
        return col;
    }

    /**
     * 写入展开后的 List 子字段数据。
     */
    private int writeExpandedListDataCells(List<?> list, ListFieldInfo listInfo, int rowIndex, int columnIndex) {
        int maxItems = list != null ? list.size() : 0;
        for (int i = 0; i < listInfo.columnSpan; i++) {
            ExportField childField = listInfo.childFields.get(i);
            for (int rowOffset = 0; rowOffset < maxItems; rowOffset++) {
                String value = "";
                if (rowOffset < list.size()) {
                    Object item = list.get(rowOffset);
                    if (item != null) {
                        value = support.getListItemFieldValue(item, childField);
                    }
                }
                writeCell(rowIndex + rowOffset, columnIndex + i, value, listInfo.parentField.getExcel());
            }
        }
        return columnIndex + listInfo.columnSpan;
    }

    /**
     * 写入普通字段数据单元格。
     */
    private int writeScalarDataCell(int rowIndex, int columnIndex, Object data, ExportField field) {
        String value = support.getFieldValue(data, field);
        writeCell(rowIndex, columnIndex, value, field.getExcel());
        return columnIndex + 1;
    }

    /**
     * 写入单个数据单元格并应用样式。
     */
    private void writeCell(int rowIndex, int columnIndex, String value,
                           com.flexibleexcel.annotation.Excel excel) {
        XSSFCell cell = support.ensureRowAndGetCell(sheet, rowIndex, columnIndex);
        cell.setCellValue(value);
        support.applyCellStyle(cell, excel);
    }

    /**
     * 连续写入指定数量的空白单元格。
     */
    private int writeBlankCells(int rowIndex, int startColumn, int count,
                                com.flexibleexcel.annotation.Excel excel) {
        int column = startColumn;
        for (int i = 0; i < count; i++) {
            writeCell(rowIndex, column, "", excel);
            column++;
        }
        return column;
    }
}
