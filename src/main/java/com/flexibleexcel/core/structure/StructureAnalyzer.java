package com.flexibleexcel.core.structure;

import com.flexibleexcel.annotation.Excel;
import com.flexibleexcel.annotation.ExcelIgnore;
import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.core.ExcelContext;

import java.util.ArrayList;
import java.util.List;

/**
 * 导出结构分析器。
 * 负责收集嵌套字段、列表字段元数据，并计算多级表头深度。
 */
public class StructureAnalyzer {

    /**
     * 收集所有需要横向展开的嵌套字段信息。
     */
    public List<NestedProcessor.NestedFieldInfo> collectNestedFieldInfos(ExcelContext context) {
        List<NestedProcessor.NestedFieldInfo> infos = new ArrayList<>();
        int currentColumn = 0;

        for (ExportField field : context.getExportFields()) {
            if (field.isNested() && field.getExcel().nested() == Excel.NestedMode.HORIZONTAL) {
                NestedProcessor.NestedFieldInfo info = NestedProcessor.buildNestedFieldInfo(field, currentColumn);
                addAllNestedInfos(infos, info);
                currentColumn += info.columnSpan;
            } else {
                currentColumn++;
            }
        }

        return infos;
    }

    /**
     * 递归拉平全部嵌套结构，便于后续统一查询。
     */
    private void addAllNestedInfos(List<NestedProcessor.NestedFieldInfo> infos, NestedProcessor.NestedFieldInfo info) {
        infos.add(info);
        if (info.columns != null) {
            for (NestedProcessor.NestedFieldInfo.NestedColumn col : info.columns) {
                if (col.isNested()) {
                    addAllNestedInfos(infos, col.nestedInfo);
                }
            }
        }
    }

    /**
     * 收集所有需要展开的 List 字段信息。
     */
    public List<ListFieldInfo> collectListFieldInfos(ExcelContext context) {
        List<ListFieldInfo> infos = new ArrayList<>();

        for (ExportField field : context.getExportFields()) {
            if (field.isList()) {
                ListFieldInfo info = buildListFieldInfo(field);
                if (info != null && info.columnSpan > 0) {
                    infos.add(info);
                }
            }
        }

        return infos;
    }

    /**
     * 构建单个 List 字段的结构信息。
     */
    private ListFieldInfo buildListFieldInfo(ExportField field) {
        ListFieldInfo info = new ListFieldInfo(field);

        Class<?> itemType = getGenericType(field);
        if (itemType == null || isSimpleJavaType(itemType)) {
            return null;
        }

        java.lang.reflect.Field[] fields = itemType.getDeclaredFields();
        for (java.lang.reflect.Field f : fields) {
            Excel excel = f.getAnnotation(Excel.class);
            if (excel != null) {
                ExcelIgnore ignore = f.getAnnotation(ExcelIgnore.class);
                if (ignore != null) {
                    continue;
                }

                ExportField childField = new ExportField(f, excel, info.childFields.size());
                info.childFields.add(childField);
                info.childHeaders.add(getFieldHeader(childField));
            }
        }

        info.columnSpan = info.childFields.size();
        if (info.columnSpan <= 1) {
            return null;
        }
        return info;
    }

    /**
     * 计算当前导出的最大表头深度。
     */
    public int calculateMaxHeaderDepth(ExcelContext context,
                                       List<NestedProcessor.NestedFieldInfo> nestedInfos,
                                       List<ListFieldInfo> listInfos) {
        int depth = 1;
        for (ExportField field : context.getExportFields()) {
            depth = Math.max(depth, calculateFieldDepth(field, 1, nestedInfos, listInfos));
        }
        return depth;
    }

    /**
     * 查找字段对应的嵌套结构信息。
     */
    NestedProcessor.NestedFieldInfo getNestedFieldInfo(ExportField field,
                                                       List<NestedProcessor.NestedFieldInfo> infos) {
        for (NestedProcessor.NestedFieldInfo info : infos) {
            if (info.parentField == field) {
                return info;
            }
        }
        return null;
    }

    /**
     * 查找字段对应的 List 结构信息。
     */
    ListFieldInfo getListFieldInfo(ExportField field, List<ListFieldInfo> infos) {
        if (infos == null) {
            return null;
        }
        for (ListFieldInfo info : infos) {
            if (info.parentField == field) {
                return info;
            }
        }
        return null;
    }

    /**
     * 计算单个字段占用的表头深度。
     */
    private int calculateFieldDepth(ExportField field, int currentDepth,
                                    List<NestedProcessor.NestedFieldInfo> nestedInfos,
                                    List<ListFieldInfo> listInfos) {
        int depth = currentDepth;

        if (field.isMap()) {
            return Math.max(depth, 2);
        }

        if (field.isNested()) {
            NestedProcessor.NestedFieldInfo nestedInfo = getNestedFieldInfo(field, nestedInfos);
            if (nestedInfo != null && nestedInfo.columnSpan > 0) {
                for (ExportField childField : nestedInfo.childFields) {
                    depth = Math.max(depth, calculateChildFieldDepth(childField, currentDepth + 1, nestedInfos, listInfos));
                }
            }
            return depth;
        }

        if (field.isList()) {
            ListFieldInfo listInfo = getListFieldInfo(field, listInfos);
            if (listInfo != null && listInfo.columnSpan > 0) {
                for (ExportField childField : listInfo.childFields) {
                    depth = Math.max(depth, calculateChildFieldDepth(childField, currentDepth + 1, nestedInfos, listInfos));
                }
                return depth;
            }
            return Math.max(depth, 2);
        }

        return depth;
    }

    /**
     * 计算子字段占用的表头深度。
     */
    private int calculateChildFieldDepth(ExportField field, int currentDepth,
                                         List<NestedProcessor.NestedFieldInfo> nestedInfos,
                                         List<ListFieldInfo> listInfos) {
        return calculateFieldDepth(field, currentDepth, nestedInfos, listInfos);
    }

    /**
     * 获取字段表头文本。
     */
    private String getFieldHeader(ExportField field) {
        if (field.getExcel() != null && !field.getExcel().header().isEmpty()) {
            return field.getExcel().header();
        }
        return field.getField().getName();
    }

    /**
     * 判断是否为简单 Java 类型。
     */
    private boolean isSimpleJavaType(Class<?> type) {
        return type.isPrimitive()
                || type == String.class
                || type == Boolean.class
                || type == Character.class
                || type == Byte.class
                || type == Short.class
                || type == Integer.class
                || type == Long.class
                || type == Float.class
                || type == Double.class
                || type == java.math.BigInteger.class
                || type == java.math.BigDecimal.class
                || Number.class.isAssignableFrom(type)
                || CharSequence.class.isAssignableFrom(type)
                || java.util.Date.class.isAssignableFrom(type)
                || java.time.LocalDate.class.isAssignableFrom(type)
                || java.time.LocalDateTime.class.isAssignableFrom(type);
    }

    /**
     * 获取 List 字段的泛型元素类型。
     */
    private Class<?> getGenericType(ExportField field) {
        try {
            java.lang.reflect.Type type = field.getField().getGenericType();
            if (type instanceof java.lang.reflect.ParameterizedType) {
                java.lang.reflect.ParameterizedType pt = (java.lang.reflect.ParameterizedType) type;
                java.lang.reflect.Type[] typeArgs = pt.getActualTypeArguments();
                if (typeArgs.length > 0) {
                    return (Class<?>) typeArgs[0];
                }
            }
        } catch (Exception e) {
            // ignore
        }
        return null;
    }
}
