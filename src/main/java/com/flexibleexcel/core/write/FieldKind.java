package com.flexibleexcel.core.write;

import com.flexibleexcel.core.ExcelContext.ExportField;

/**
 * 导出字段的写入分类。
 * 用于统一三类写入器中的字段分派逻辑。
 */
enum FieldKind {
    MAP,
    NESTED,
    LIST,
    SCALAR;

    /**
     * 根据字段特征推导写入分类。
     */
    static FieldKind from(ExportField field) {
        if (field.isMap()) {
            return MAP;
        }
        if (field.isNested()) {
            return NESTED;
        }
        if (field.isList()) {
            return LIST;
        }
        return SCALAR;
    }
}
