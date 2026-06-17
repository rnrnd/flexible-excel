package com.flexibleexcel.core.write;

import com.flexibleexcel.annotation.ExcelConfig;
import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.core.ExportState;
import com.flexibleexcel.core.ExportSupport;
import com.flexibleexcel.core.structure.ListFieldInfo;
import com.flexibleexcel.core.structure.MapKeyCollector;
import com.flexibleexcel.core.structure.NestedProcessor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

/**
 * 表头写入协作对象，负责组织简单表头和复杂表头的写入流程。
 */
public class HeaderWriter extends AbstractFieldWriter {
    private final ExportSupport support;
    private final XSSFSheet sheet;
    private final List<?> dataList;
    private final ExcelConfig config;
    private final ExportState state;

    public HeaderWriter(ExportSupport support, XSSFSheet sheet, List<?> dataList,
                        ExcelConfig config, ExportState state) {
        super(state);
        this.support = support;
        this.sheet = sheet;
        this.dataList = dataList;
        this.config = config;
        this.state = state;
    }

    /**
     * 执行表头写入。
     */
    public void write() {
        if (state.maxHeaderDepth > 1) {
            writeNestedHeader();
        } else {
            writeSimpleHeader();
        }
    }

    /**
     * 写入简单表头。
     */
    private void writeSimpleHeader() {
        XSSFRow headerRow = sheet.createRow(0);
        forEachField(0, (field, currentColumn, kind) -> {
            switch (kind) {
                case MAP:
                    return writeSimpleMapHeaders(headerRow, currentColumn, field);
                case LIST:
                    return writeSimpleListHeaders(headerRow, currentColumn, field);
                default:
                    return writeSimpleValueHeader(headerRow, currentColumn, field);
            }
        });

        if (config != null) {
            support.applyHeaderStyleToRow(sheet.getRow(0), config);
        }
    }

    /**
     * 写入多行复杂表头。
     */
    private void writeNestedHeader() {
        XSSFRow[] headerRows = new XSSFRow[state.maxHeaderDepth];
        for (int i = 0; i < state.maxHeaderDepth; i++) {
            headerRows[i] = sheet.createRow(i);
        }

        writeHeaderFields(headerRows, 0);

        for (int i = 0; i < state.maxHeaderDepth; i++) {
            if (i == 0) {
                support.applyHeaderStyleToRow(headerRows[i], config);
            } else {
                support.applySubHeaderStyleToRow(headerRows[i]);
            }
        }
    }

    /**
     * 递归写入复杂表头字段。
     *
     * @param headerRows 表头行数组
     * @param currentDepth 当前深度
     */
    private void writeHeaderFields(XSSFRow[] headerRows, int currentDepth) {
        forEachField(0, (field, currentColumn, kind) -> {
            switch (kind) {
                case MAP:
                    return writeMapHeaderFields(headerRows, field, currentDepth, currentColumn);
                case NESTED:
                    return writeNestedHeaderField(headerRows, field, currentDepth, currentColumn);
                case LIST:
                    return writeListHeaderField(headerRows, field, currentDepth, currentColumn);
                default:
                    return writeScalarHeaderField(headerRows, field, currentDepth, currentColumn);
            }
        });
    }

    /**
     * 写入简单表头中的 Map 列。
     */
    private int writeSimpleMapHeaders(XSSFRow headerRow, int currentColumn, ExportField field) {
        List<Object> allKeys = MapKeyCollector.getAllKeys(dataList, field);
        for (Object key : allKeys) {
            XSSFCell cell = headerRow.createCell(currentColumn);
            cell.setCellValue(key != null ? key.toString() : "");
            support.applyHeaderStyle(cell);
            currentColumn++;
        }
        return currentColumn;
    }

    /**
     * 写入简单表头中的 List 列。
     */
    private int writeSimpleListHeaders(XSSFRow headerRow, int currentColumn, ExportField field) {
        ListFieldInfo listInfo = state.getListFieldInfo(field);
        if (listInfo != null && listInfo.columnSpan > 0) {
            for (int i = 0; i < listInfo.columnSpan; i++) {
                XSSFCell cell = headerRow.createCell(currentColumn);
                if (i < listInfo.childHeaders.size()) {
                    cell.setCellValue(listInfo.childHeaders.get(i));
                } else {
                    cell.setCellValue("");
                }
                support.applyHeaderStyle(cell);
                currentColumn++;
            }
            return currentColumn;
        }
        return writeSimpleValueHeader(headerRow, currentColumn, field);
    }

    /**
     * 写入简单表头中的普通列。
     */
    private int writeSimpleValueHeader(XSSFRow headerRow, int currentColumn, ExportField field) {
        XSSFCell cell = headerRow.createCell(currentColumn);
        cell.setCellValue(support.getHeaderText(field));
        support.applyHeaderStyle(cell);
        return currentColumn + 1;
    }

    /**
     * 写入复杂表头中的 Map 字段。
     */
    private int writeMapHeaderFields(XSSFRow[] headerRows, ExportField field,
                                     int currentDepth, int currentColumn) {
        List<Object> allKeys = MapKeyCollector.getAllKeys(dataList, field);
        int columnSpan = allKeys.size();

        writeMainHeaderCell(headerRows[currentDepth], currentDepth, currentColumn,
                columnSpan, support.getHeaderText(field), field.getExcel(), false);
        writeSubHeaders(headerRows, currentDepth, currentColumn, columnSpan,
                index -> {
                    Object key = allKeys.get(index);
                    return key != null ? key.toString() : "";
                });

        return currentColumn + columnSpan;
    }

    /**
     * 写入复杂表头中的嵌套字段。
     */
    private int writeNestedHeaderField(XSSFRow[] headerRows, ExportField field,
                                       int currentDepth, int currentColumn) {
        NestedProcessor.NestedFieldInfo nestedInfo = state.getNestedFieldInfo(field);
        if (nestedInfo != null) {
            return writeNestedHeaderRecursive(headerRows, nestedInfo, currentDepth, currentColumn);
        }
        return writeScalarHeaderField(headerRows, field, currentDepth, currentColumn);
    }

    /**
     * 写入复杂表头中的 List 字段。
     */
    private int writeListHeaderField(XSSFRow[] headerRows, ExportField field,
                                     int currentDepth, int currentColumn) {
        ListFieldInfo listInfo = state.getListFieldInfo(field);
        if (listInfo == null || listInfo.columnSpan <= 0) {
            return writeScalarHeaderField(headerRows, field, currentDepth, currentColumn);
        }

        writeMainHeaderCell(headerRows[currentDepth], currentDepth, currentColumn,
                listInfo.columnSpan, support.getHeaderText(field), field.getExcel(), true);
        writeSubHeaders(headerRows, currentDepth, currentColumn, listInfo.columnSpan,
                index -> index < listInfo.childHeaders.size() ? listInfo.childHeaders.get(index) : "");

        return currentColumn + listInfo.columnSpan;
    }

    /**
     * 写入复杂表头中的普通字段。
     */
    private int writeScalarHeaderField(XSSFRow[] headerRows, ExportField field,
                                       int currentDepth, int currentColumn) {
        XSSFCell mainCell = headerRows[currentDepth].createCell(currentColumn);
        mainCell.setCellValue(support.getHeaderText(field));
        support.applyNestedMainHeaderStyle(mainCell, config, field.getExcel());

        int rowSpan = state.maxHeaderDepth - currentDepth;
        if (rowSpan > 1) {
            CellRangeAddress mergeRegion = new CellRangeAddress(
                    currentDepth, state.maxHeaderDepth - 1, currentColumn, currentColumn);
            sheet.addMergedRegion(mergeRegion);
            support.applyHeaderBordersToMergedRegion(sheet, mergeRegion);
        }

        return currentColumn + 1;
    }

    /**
     * 递归写入复杂嵌套表头。
     */
    private int writeNestedHeaderRecursive(XSSFRow[] headerRows,
                                           NestedProcessor.NestedFieldInfo nestedInfo,
                                           int currentDepth, int currentColumn) {
        writeMainHeaderCell(headerRows[currentDepth], currentDepth, currentColumn,
                nestedInfo.columnSpan, nestedInfo.parentHeader, nestedInfo.parentField.getExcel(), true);

        if (currentDepth + 1 < state.maxHeaderDepth && nestedInfo.columns != null) {
            int childCol = currentColumn;
            for (NestedProcessor.NestedFieldInfo.NestedColumn col : nestedInfo.columns) {
                if (col.isNested()) {
                    childCol = writeNestedHeaderRecursive(headerRows, col.nestedInfo, currentDepth + 1, childCol);
                } else {
                    XSSFCell subCell = headerRows[currentDepth + 1].createCell(childCol);
                    subCell.setCellValue(col.header);
                    support.applyNestedSubHeaderStyle(subCell);

                    if (currentDepth + 1 < state.maxHeaderDepth - 1) {
                        CellRangeAddress mergeRegion = new CellRangeAddress(
                                currentDepth + 1, state.maxHeaderDepth - 1, childCol, childCol);
                        sheet.addMergedRegion(mergeRegion);
                        support.applyHeaderBordersToMergedRegion(sheet, mergeRegion);
                    }
                    childCol++;
                }
            }
        }

        return currentColumn + nestedInfo.columnSpan;
    }

    /**
     * 写入主表头单元格，并在需要时执行横向合并。
     */
    private void writeMainHeaderCell(XSSFRow row, int currentDepth, int currentColumn,
                                     int columnSpan, String text,
                                     com.flexibleexcel.annotation.Excel excel,
                                     boolean applyMergeBorders) {
        XSSFCell mainCell = row.createCell(currentColumn);
        mainCell.setCellValue(text);
        support.applyNestedMainHeaderStyle(mainCell, config, excel);

        if (columnSpan > 1) {
            CellRangeAddress mergeRegion = new CellRangeAddress(
                    currentDepth, currentDepth, currentColumn, currentColumn + columnSpan - 1);
            sheet.addMergedRegion(mergeRegion);
            if (applyMergeBorders) {
                support.applyHeaderBordersToMergedRegion(sheet, mergeRegion);
            }
        }
    }

    /**
     * 批量写入次级表头单元格。
     */
    private void writeSubHeaders(XSSFRow[] headerRows, int currentDepth, int currentColumn,
                                 int columnSpan, HeaderTextProvider textProvider) {
        if (currentDepth + 1 >= state.maxHeaderDepth) {
            return;
        }

        for (int i = 0; i < columnSpan; i++) {
            XSSFCell subCell = headerRows[currentDepth + 1].createCell(currentColumn + i);
            subCell.setCellValue(textProvider.get(i));
            support.applyNestedSubHeaderStyle(subCell);
        }
    }

    /**
     * 次级表头文本提供器。
     */
    private interface HeaderTextProvider {
        String get(int index);
    }
}
