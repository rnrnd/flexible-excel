package com.flexibleexcel.core.write;

import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.core.ExportState;

/**
 * 字段遍历写入模板基类。
 * 负责统一按字段顺序遍历并推进列游标。
 */
abstract class AbstractFieldWriter {
    protected final ExportState state;

    protected AbstractFieldWriter(ExportState state) {
        this.state = state;
    }

    /**
     * 遍历当前导出上下文中的所有字段，并累计列位置。
     */
    protected int forEachField(int startColumn, FieldColumnHandler handler) {
        int currentColumn = startColumn;
        for (ExportField field : state.context.getExportFields()) {
            currentColumn = handler.handle(field, currentColumn, FieldKind.from(field));
        }
        return currentColumn;
    }

    /**
     * 字段列处理函数。
     */
    @FunctionalInterface
    protected interface FieldColumnHandler {
        int handle(ExportField field, int currentColumn, FieldKind kind);
    }
}
