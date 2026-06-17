package com.flexibleexcel.core.structure;

import com.flexibleexcel.core.ExcelContext.ExportField;

import java.util.ArrayList;
import java.util.List;

/**
 * List 字段的结构信息。
 * 用于描述列表子字段展开后的列定义和表头信息。
 */
public class ListFieldInfo {
    public ExportField parentField;
    public List<ExportField> childFields;
    public List<String> childHeaders;
    public int columnSpan;

    public ListFieldInfo(ExportField parentField) {
        this.parentField = parentField;
        this.childFields = new ArrayList<>();
        this.childHeaders = new ArrayList<>();
    }
}
