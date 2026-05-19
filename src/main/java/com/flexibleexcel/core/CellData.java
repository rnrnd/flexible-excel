package com.flexibleexcel.core;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 单元格数据
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class CellData {

    /**
     * 单元格值
     */
    private Object value;

    /**
     * 起始行索引
     */
    private int startRow;

    /**
     * 结束行索引
     */
    private int endRow;

    /**
     * 起始列索引
     */
    private int startColumn;

    /**
     * 结束列索引
     */
    private int endColumn;

    /**
     * 单元格样式
     */
    private CellStyle style;

    /**
     * 是否需要合并
     */
    private boolean needMerge;

    /**
     * 表头文本（用于Map展开时）
     */
    private String headerText;

    /**
     * 是否为表头单元格
     */
    private boolean isHeader;

    /**
     * 创建简单的单元格数据
     */
    public static CellData of(Object value, int row, int column) {
        return CellData.builder()
                .value(value)
                .startRow(row)
                .endRow(row)
                .startColumn(column)
                .endColumn(column)
                .style(CellStyle.defaultStyle())
                .needMerge(false)
                .build();
    }

    /**
     * 创建需要合并的单元格数据
     */
    public static CellData merged(Object value, int startRow, int endRow, int startColumn, int endColumn) {
        return CellData.builder()
                .value(value)
                .startRow(startRow)
                .endRow(endRow)
                .startColumn(startColumn)
                .endColumn(endColumn)
                .style(CellStyle.defaultStyle())
                .needMerge(startRow != endRow || startColumn != endColumn)
                .build();
    }

    /**
     * 创建表头单元格
     */
    public static CellData header(String text, int row, int column) {
        return CellData.builder()
                .value(text)
                .startRow(row)
                .endRow(row)
                .startColumn(column)
                .endColumn(column)
                .isHeader(true)
                .style(CellStyle.defaultStyle())
                .needMerge(false)
                .build();
    }

    /**
     * 创建需要合并的表头单元格
     */
    public static CellData mergedHeader(String text, int startRow, int endRow, int startColumn, int endColumn) {
        return CellData.builder()
                .value(text)
                .startRow(startRow)
                .endRow(endRow)
                .startColumn(startColumn)
                .endColumn(endColumn)
                .isHeader(true)
                .style(CellStyle.defaultStyle())
                .needMerge(true)
                .build();
    }
}
