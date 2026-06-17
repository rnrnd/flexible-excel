package com.flexibleexcel.core.write;

import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.core.ExportState;
import com.flexibleexcel.core.ExportSupport;
import com.flexibleexcel.core.structure.ListFieldInfo;
import com.flexibleexcel.core.structure.MapKeyCollector;
import com.flexibleexcel.core.structure.NestedProcessor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

/**
 * 合并策略协作对象，负责数据区纵向合并流程。
 */
public class MergeHandler {
    private final ExportSupport support;
    private final XSSFSheet sheet;
    private final List<?> dataList;
    private final ExportState state;

    public MergeHandler(ExportSupport support, XSSFSheet sheet, List<?> dataList, ExportState state) {
        this.support = support;
        this.sheet = sheet;
        this.dataList = dataList;
        this.state = state;
    }

    /**
     * 执行数据区合并。
     */
    public void apply() {
        int currentRow = state.context.getDataStartRow();

        for (Object data : dataList) {
            int maxListSize = support.getMaxListSize(data, state);
            int currentColumn = 0;

            for (ExportField field : state.context.getExportFields()) {
                switch (FieldKind.from(field)) {
                    case MAP:
                        currentColumn = applyMapMerges(field, currentRow, maxListSize, currentColumn);
                        break;
                    case NESTED:
                        currentColumn = applyNestedOrScalarMerges(field, currentRow, maxListSize, currentColumn);
                        break;
                    case LIST:
                        currentColumn = skipListMergeColumns(field, currentColumn);
                        break;
                    default:
                        currentColumn = applyScalarMerge(currentRow, maxListSize, currentColumn);
                        break;
                }
            }

            currentRow += Math.max(maxListSize, 1);
        }
    }

    /**
     * 对 Map 字段对应列执行纵向合并。
     */
    private int applyMapMerges(ExportField field, int rowIndex, int rowSpan, int columnIndex) {
        List<Object> allKeys = MapKeyCollector.getAllKeys(dataList, field);
        for (int i = 0; i < allKeys.size(); i++) {
            applyVerticalMergeIfNeeded(rowIndex, rowSpan, columnIndex);
            columnIndex = nextColumn(columnIndex);
        }
        return columnIndex;
    }

    /**
     * 对嵌套字段或普通字段执行纵向合并。
     */
    private int applyNestedOrScalarMerges(ExportField field, int rowIndex, int rowSpan, int columnIndex) {
        NestedProcessor.NestedFieldInfo nestedInfo = state.getNestedFieldInfo(field);
        if (nestedInfo != null) {
            return applyNestedMergesRecursive(nestedInfo, rowIndex, rowSpan, columnIndex);
        }
        return applyScalarMerge(rowIndex, rowSpan, columnIndex);
    }

    /**
     * 跳过 List 字段占用的列数。
     */
    private int skipListMergeColumns(ExportField field, int columnIndex) {
        ListFieldInfo listInfo = state.getListFieldInfo(field);
        if (listInfo != null && listInfo.columnSpan > 0) {
            return advanceColumns(columnIndex, listInfo.columnSpan);
        }
        return nextColumn(columnIndex);
    }

    /**
     * 对普通列执行纵向合并。
     */
    private int applyScalarMerge(int rowIndex, int rowSpan, int columnIndex) {
        applyVerticalMergeIfNeeded(rowIndex, rowSpan, columnIndex);
        return nextColumn(columnIndex);
    }

    /**
     * 当存在多行数据时执行单列纵向合并。
     */
    private void applyVerticalMergeIfNeeded(int rowIndex, int rowSpan, int columnIndex) {
        if (rowSpan > 1) {
            CellRangeAddress mergeRegion = new CellRangeAddress(
                    rowIndex, rowIndex + rowSpan - 1, columnIndex, columnIndex
            );
            sheet.addMergedRegion(mergeRegion);
            support.applyBordersToMergedRegion(sheet, mergeRegion);
        }
    }

    /**
     * 递归应用嵌套字段的纵向合并。
     */
    private int applyNestedMergesRecursive(NestedProcessor.NestedFieldInfo nestedInfo,
                                           int currentRow, int maxListSize, int currentColumn) {
        int col = currentColumn;
        for (NestedProcessor.NestedFieldInfo.NestedColumn colDef : nestedInfo.columns) {
            if (colDef.isNested()) {
                col = applyNestedMergesRecursive(colDef.nestedInfo, currentRow, maxListSize, col);
            } else {
                applyVerticalMergeIfNeeded(currentRow, maxListSize, col);
                col = nextColumn(col);
            }
        }
        return col;
    }

    /**
     * 推进一个列位置。
     */
    private int nextColumn(int columnIndex) {
        return columnIndex + 1;
    }

    /**
     * 推进多个列位置。
     */
    private int advanceColumns(int columnIndex, int span) {
        return columnIndex + span;
    }
}
