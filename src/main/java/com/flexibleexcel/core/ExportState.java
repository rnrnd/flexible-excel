package com.flexibleexcel.core;

import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.core.structure.ListFieldInfo;
import com.flexibleexcel.core.structure.NestedProcessor;

import java.util.List;

/**
 * 单次导出过程中的阶段状态。
 * 统一持有本次导出的上下文、结构元数据以及表头深度信息。
 */
public class ExportState {
    public final ExcelContext context;
    public final List<NestedProcessor.NestedFieldInfo> nestedFieldInfos;
    public final List<ListFieldInfo> listFieldInfos;
    public final int maxHeaderDepth;

    public ExportState(ExcelContext context,
                       List<NestedProcessor.NestedFieldInfo> nestedFieldInfos,
                       List<ListFieldInfo> listFieldInfos,
                       int maxHeaderDepth) {
        this.context = context;
        this.nestedFieldInfos = nestedFieldInfos;
        this.listFieldInfos = listFieldInfos;
        this.maxHeaderDepth = maxHeaderDepth;
    }

    /**
     * 根据字段查找对应的嵌套结构信息。
     */
    public NestedProcessor.NestedFieldInfo getNestedFieldInfo(ExportField field) {
        for (NestedProcessor.NestedFieldInfo info : nestedFieldInfos) {
            if (info.parentField == field) {
                return info;
            }
        }
        return null;
    }

    /**
     * 根据字段查找对应的列表结构信息。
     */
    public ListFieldInfo getListFieldInfo(ExportField field) {
        if (listFieldInfos == null) {
            return null;
        }
        for (ListFieldInfo info : listFieldInfos) {
            if (info.parentField == field) {
                return info;
            }
        }
        return null;
    }
}
