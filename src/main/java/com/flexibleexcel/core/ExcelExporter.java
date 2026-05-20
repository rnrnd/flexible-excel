package com.flexibleexcel.core;

import com.flexibleexcel.annotation.Excel;
import com.flexibleexcel.annotation.ExcelConfig;
import com.flexibleexcel.annotation.ExcelIgnore;
import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.exception.ExcelExportException;
import com.flexibleexcel.processor.MapProcessor;
import com.flexibleexcel.processor.NestedProcessor;
import com.flexibleexcel.util.ReflectionUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.OutputStream;
import java.nio.file.Files;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

/**
 * Excel导出器
 * 核心导出类
 */
public class ExcelExporter {

    /**
     * List类型字段信息（用于展开次级表头）
     */
    private static class ListFieldInfo {
        ExportField parentField;
        List<ExportField> childFields;
        List<String> childHeaders;
        int columnSpan;

        public ListFieldInfo(ExportField parentField) {
            this.parentField = parentField;
            this.childFields = new ArrayList<>();
            this.childHeaders = new ArrayList<>();
        }
    }

    private ExcelContext context;
    /**
     * 嵌套字段信息列表
     */
    private List<NestedProcessor.NestedFieldInfo> nestedFieldInfos;

    /**
     * List类型字段信息列表（用于展开次级表头）
     */
    private List<ListFieldInfo> listFieldInfos;

    /**
     * 多Sheet导出时的Sheet数据列表
     */
    private final List<SheetData> sheetDataList = new ArrayList<>();

    /**
     * 创建导出器实例
     */
    public static ExcelExporter create() {
        return new ExcelExporter();
    }

    /**
     * 添加一个Sheet（链式调用）
     *
     * @param sheetName Sheet名称
     * @param clazz     数据类（带@Excel注解）
     * @param dataList  数据列表
     * @return 当前导出器实例
     */
    public ExcelExporter sheet(String sheetName, Class<?> clazz, List<?> dataList) {
        this.sheetDataList.add(SheetData.of(sheetName, clazz, dataList));
        return this;
    }

    /**
     * 添加一个Sheet，使用类上的@ExcelConfig注解的sheetName（链式调用）
     *
     * @param clazz    数据类（带@Excel注解）
     * @param dataList 数据列表
     * @return 当前导出器实例
     */
    public ExcelExporter sheet(Class<?> clazz, List<?> dataList) {
        this.sheetDataList.add(SheetData.of(clazz, dataList));
        return this;
    }

    /**
     * 导出多个Sheet到文件（链式调用结束）
     *
     * @param filePath 输出文件路径
     */
    public void toFile(String filePath) {
        toFile(new File(filePath));
    }

    /**
     * 导出多个Sheet到文件（链式调用结束）
     *
     * @param file 输出文件
     */
    public void toFile(File file) {
        try (OutputStream out = Files.newOutputStream(file.toPath())) {
            toOutputStream(out);
        } catch (Exception e) {
            throw new ExcelExportException("导出Excel失败: " + e.getMessage(), e);
        }
    }

    /**
     * 导出多个Sheet到输出流（链式调用结束）
     *
     * @param out 输出流
     */
    public void toOutputStream(OutputStream out) {
        if (sheetDataList.isEmpty()) {
            throw new ExcelExportException("没有添加任何Sheet数据，请先调用sheet()方法添加数据");
        }

        try {
            XSSFWorkbook workbook = new XSSFWorkbook();

            for (SheetData sheetData : sheetDataList) {
                // 为每个Sheet创建上下文
                this.context = ExcelContext.create(sheetData.getClazz(), sheetData.getDataList());
                this.nestedFieldInfos = collectNestedFieldInfos(sheetData.getDataList());
                this.listFieldInfos = collectListFieldInfos();

                ExcelConfig config = sheetData.getClazz().getAnnotation(ExcelConfig.class);
                int headerRows = hasNestedFields() || hasListFieldsForExpand() ? 2 : 1;
                context.setHeaderRows(headerRows);
                context.setDataStartRow(headerRows);

                // 创建Sheet
                XSSFSheet sheet = createSheet(workbook, sheetData.getSheetName());

                // 写入表头
                writeHeader(sheet, sheetData.getDataList(), config);

                // 写入数据
                writeData(sheet, sheetData.getDataList());

                // 处理单元格合并
                applyMerges(sheet, sheetData.getDataList());

                // 设置列宽
                applyColumnWidths(sheet);
            }

            // 写入文件
            workbook.write(out);
            workbook.close();

        } catch (Exception e) {
            throw new ExcelExportException("导出Excel失败: " + e.getMessage(), e);
        }
    }

    /**
     * 导出到文件
     *
     * @param clazz     数据类（带@Excel注解）
     * @param dataList  数据列表
     * @param filePath  输出文件路径
     */
    public void export(Class<?> clazz, List<?> dataList, String filePath) {
        export(clazz, dataList, new File(filePath));
    }

    /**
     * 导出到文件
     *
     * @param clazz     数据类（带@Excel注解）
     * @param dataList  数据列表
     * @param file      输出文件
     */
    public void export(Class<?> clazz, List<?> dataList, File file) {
        try (OutputStream out = Files.newOutputStream(file.toPath())) {
            export(clazz, dataList, out);
        } catch (Exception e) {
            throw new ExcelExportException("导出Excel失败: " + e.getMessage(), e);
        }
    }

    /**
     * 导出到输出流
     *
     * @param clazz    数据类（带@Excel注解）
     * @param dataList 数据列表
     * @param out      输出流
     */
    public void export(Class<?> clazz, List<?> dataList, OutputStream out) {
        try {
            // 初始化上下文
            this.context = ExcelContext.create(clazz, dataList);

            // 收集嵌套字段信息
            this.nestedFieldInfos = collectNestedFieldInfos(dataList);
            this.listFieldInfos = collectListFieldInfos();

            // 获取配置
            ExcelConfig config = clazz.getAnnotation(ExcelConfig.class);

            // 确定表头行数（是否有嵌套对象或需要展开的List字段）
            int headerRows = hasNestedFields() || hasListFieldsForExpand() ? 2 : 1;
            context.setHeaderRows(headerRows);
            context.setDataStartRow(headerRows);

            // 创建工作簿
            XSSFWorkbook workbook = new XSSFWorkbook();
            String sheetName = (config != null && !config.sheetName().isEmpty()) ? config.sheetName() : "Sheet1";
            XSSFSheet sheet = createSheet(workbook, sheetName);

            // 写入表头
            writeHeader(sheet, dataList, config);

            // 写入数据
            writeData(sheet, dataList);

            // 处理单元格合并
            applyMerges(sheet, dataList);

            // 设置列宽
            applyColumnWidths(sheet);

            // 写入文件
            workbook.write(out);
            workbook.close();

        } catch (Exception e) {
            throw new ExcelExportException("导出Excel失败: " + e.getMessage(), e);
        }
    }

    /**
     * 收集嵌套字段信息
     */
    private List<NestedProcessor.NestedFieldInfo> collectNestedFieldInfos(List<?> dataList) {
        List<NestedProcessor.NestedFieldInfo> infos = new ArrayList<>();
        int currentColumn = 0;

        for (ExportField field : context.getExportFields()) {
            if (field.isNested() && field.getExcel().nested() == Excel.NestedMode.HORIZONTAL) {
                NestedProcessor.NestedFieldInfo info = NestedProcessor.buildNestedFieldInfo(field, currentColumn, dataList);
                infos.add(info);
                currentColumn += info.columnSpan;
            } else {
                currentColumn++;
            }
        }

        return infos;
    }

    /**
     * 收集List类型字段信息（用于展开次级表头）
     */
    private List<ListFieldInfo> collectListFieldInfos() {
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
     * 构建List字段信息
     */
    private ListFieldInfo buildListFieldInfo(ExportField field) {
        ListFieldInfo info = new ListFieldInfo(field);

        // 获取List的泛型类型
        Class<?> itemType = getGenericType(field);
        if (itemType == null || itemType == String.class || itemType.isPrimitive()) {
            // 简单类型（如List<String>），不展开次级表头
            return null;
        }

        // 遍历itemType中所有带@Excel注解的字段
        java.lang.reflect.Field[] fields = itemType.getDeclaredFields();
        for (java.lang.reflect.Field f : fields) {
            Excel excel = f.getAnnotation(Excel.class);
            if (excel != null) {
                // 排除@ExcelIgnore
                ExcelIgnore ignore = f.getAnnotation(ExcelIgnore.class);
                if (ignore != null) {
                    continue;
                }

                // 创建子字段信息
                ExportField childField = new ExportField(f, excel, info.childFields.size());
                info.childFields.add(childField);
                info.childHeaders.add(getFieldHeader(childField));
            }
        }

        info.columnSpan = info.childFields.size();

        // 如果只有一个字段且名称简单，可以不展开
        if (info.columnSpan <= 1) {
            return null;
        }

        return info;
    }

    /**
     * 获取List字段的泛型类型
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

    /**
     * 获取字段的表头文本
     */
    private String getFieldHeader(ExportField field) {
        if (field.getExcel() != null && !field.getExcel().header().isEmpty()) {
            return field.getExcel().header();
        }
        return field.getField().getName();
    }

    /**
     * 检查是否有嵌套字段
     */
    private boolean hasNestedFields() {
        for (ExportField field : context.getExportFields()) {
            if (field.isNested()) {
                return true;
            }
        }
        return false;
    }

    /**
     * 检查是否有需要展开次级表头的List字段
     */
    private boolean hasListFieldsForExpand() {
        return listFieldInfos != null && !listFieldInfos.isEmpty();
    }

    /**
     * 获取嵌套字段信息
     */
    private NestedProcessor.NestedFieldInfo getNestedFieldInfo(ExportField field) {
        for (NestedProcessor.NestedFieldInfo info : nestedFieldInfos) {
            if (info.parentField == field) {
                return info;
            }
        }
        return null;
    }

    /**
     * 获取List字段信息
     */
    private ListFieldInfo getListFieldInfo(ExportField field) {
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

    /**
     * 创建工作表
     */
    private XSSFSheet createSheet(XSSFWorkbook workbook, String sheetName) {
        if (sheetName == null || sheetName.isEmpty()) {
            sheetName = "Sheet1";
        }
        return workbook.createSheet(sheetName);
    }

    /**
     * 写入表头
     */
    private void writeHeader(XSSFSheet sheet, List<?> dataList, ExcelConfig config) {
        if (hasNestedFields() || hasListFieldsForExpand()) {
            writeNestedHeader(sheet, dataList, config);
        } else {
            writeSimpleHeader(sheet, dataList, config);
        }
    }

    /**
     * 写入简单表头（无嵌套）
     */
    private void writeSimpleHeader(XSSFSheet sheet, List<?> dataList, ExcelConfig config) {
        XSSFRow headerRow = sheet.createRow(0);
        int currentColumn = 0;

        for (ExportField field : context.getExportFields()) {
            if (field.isMap()) {
                // Map类型：获取所有key作为次级表头
                List<Object> allKeys = MapProcessor.getAllKeys(dataList, field);
                for (Object key : allKeys) {
                    XSSFCell cell = headerRow.createCell(currentColumn);
                    cell.setCellValue(key != null ? key.toString() : "");
                    applyHeaderStyle(cell);
                    currentColumn++;
                }
            } else if (field.isList()) {
                // List类型：检查是否需要展开次级表头
                ListFieldInfo listInfo = getListFieldInfo(field);
                if (listInfo != null && listInfo.columnSpan > 0) {
                    // 展开List的子字段
                    for (int i = 0; i < listInfo.columnSpan; i++) {
                        XSSFCell cell = headerRow.createCell(currentColumn);
                        if (i < listInfo.childHeaders.size()) {
                            cell.setCellValue(listInfo.childHeaders.get(i));
                        } else {
                            cell.setCellValue("");
                        }
                        applyHeaderStyle(cell);
                        currentColumn++;
                    }
                } else {
                    // 不展开，直接显示表头
                    XSSFCell cell = headerRow.createCell(currentColumn);
                    cell.setCellValue(getHeaderText(field));
                    applyHeaderStyle(cell);
                    currentColumn++;
                }
            } else {
                // 普通字段
                XSSFCell cell = headerRow.createCell(currentColumn);
                cell.setCellValue(getHeaderText(field));
                applyHeaderStyle(cell);
                currentColumn++;
            }
        }

        // 应用表头样式到整行
        if (config != null) {
            applyHeaderStyleToRow(sheet.getRow(0), config);
        }
    }

    /**
     * 写入嵌套表头（两级表头）
     */
    private void writeNestedHeader(XSSFSheet sheet, List<?> dataList, ExcelConfig config) {
        // 创建主表头行（第0行）和次级表头行（第1行）
        XSSFRow mainHeaderRow = sheet.createRow(0);
        XSSFRow subHeaderRow = sheet.createRow(1);
        int currentColumn = 0;

        for (ExportField field : context.getExportFields()) {
            if (field.isMap()) {
                // Map类型
                List<Object> allKeys = MapProcessor.getAllKeys(dataList, field);

                // 主表头合并
                if (allKeys.size() > 1) {
                    CellRangeAddress mergeRegion = new CellRangeAddress(0, 0, currentColumn, currentColumn + allKeys.size() - 1);
                    sheet.addMergedRegion(mergeRegion);
                }

                XSSFCell mainCell = mainHeaderRow.createCell(currentColumn);
                mainCell.setCellValue(getHeaderText(field));
                applyNestedMainHeaderStyle(mainCell, config, field.getExcel());

                // 次级表头
                for (Object key : allKeys) {
                    XSSFCell subCell = subHeaderRow.createCell(currentColumn);
                    subCell.setCellValue(key != null ? key.toString() : "");
                    applyNestedSubHeaderStyle(subCell);
                    currentColumn++;
                }
            } else if (field.isNested()) {
                // 嵌套对象类型
                NestedProcessor.NestedFieldInfo nestedInfo = getNestedFieldInfo(field);

                if (nestedInfo != null && nestedInfo.columnSpan > 1 && field.getExcel().mergeNestedHeader()) {
                    // 主表头合并
                    CellRangeAddress mergeRegion = new CellRangeAddress(0, 0, currentColumn, currentColumn + nestedInfo.columnSpan - 1);
                    sheet.addMergedRegion(mergeRegion);
                    // 为合并区域添加边框
                    applyHeaderBordersToMergedRegion(sheet, mergeRegion);
                }

                // 主表头
                XSSFCell mainCell = mainHeaderRow.createCell(currentColumn);
                mainCell.setCellValue(nestedInfo != null ? nestedInfo.parentHeader : getHeaderText(field));
                applyNestedMainHeaderStyle(mainCell, config, field.getExcel());

                // 次级表头
                if (nestedInfo != null) {
                    for (int i = 0; i < nestedInfo.columnSpan; i++) {
                        XSSFCell subCell = subHeaderRow.createCell(currentColumn + i);
                        if (i < nestedInfo.childHeaders.size()) {
                            subCell.setCellValue(nestedInfo.childHeaders.get(i));
                        } else {
                            subCell.setCellValue("");
                        }
                        applyNestedSubHeaderStyle(subCell);
                    }
                }

                currentColumn += (nestedInfo != null ? nestedInfo.columnSpan : 1);
            } else if (field.isList()) {
                // List类型 - 检查是否需要展开次级表头
                ListFieldInfo listInfo = getListFieldInfo(field);
                if (listInfo != null && listInfo.columnSpan > 0) {
                    // 需要展开
                    // 主表头合并
                    if (listInfo.columnSpan > 1) {
                        CellRangeAddress mergeRegion = new CellRangeAddress(0, 0, currentColumn, currentColumn + listInfo.columnSpan - 1);
                        sheet.addMergedRegion(mergeRegion);
                        // 为合并区域添加边框
                        applyHeaderBordersToMergedRegion(sheet, mergeRegion);
                    }

                    // 主表头
                    XSSFCell mainCell = mainHeaderRow.createCell(currentColumn);
                    mainCell.setCellValue(getHeaderText(field));
                    applyNestedMainHeaderStyle(mainCell, config, field.getExcel());

                    // 次级表头
                    for (int i = 0; i < listInfo.columnSpan; i++) {
                        XSSFCell subCell = subHeaderRow.createCell(currentColumn + i);
                        if (i < listInfo.childHeaders.size()) {
                            subCell.setCellValue(listInfo.childHeaders.get(i));
                        } else {
                            subCell.setCellValue("");
                        }
                        applyNestedSubHeaderStyle(subCell);
                    }

                    currentColumn += listInfo.columnSpan;
                } else {
                    // 不展开，主表头跨两行（合并第0行和第1行的单元格）
                    XSSFCell mainCell = mainHeaderRow.createCell(currentColumn);
                    mainCell.setCellValue(getHeaderText(field));
                    applyNestedMainHeaderStyle(mainCell, config, field.getExcel());

                    XSSFCell subCell = subHeaderRow.createCell(currentColumn);
                    subCell.setCellValue("");
                    applyNestedSubHeaderStyle(subCell);

                    // 合并主表头和次级表头的单元格
                    CellRangeAddress mergeRegion = new CellRangeAddress(0, 1, currentColumn, currentColumn);
                    sheet.addMergedRegion(mergeRegion);

                    currentColumn++;
                }
            } else {
                // 普通字段 - 主表头跨两行（合并第0行和第1行的单元格）
                XSSFCell mainCell = mainHeaderRow.createCell(currentColumn);
                mainCell.setCellValue(getHeaderText(field));
                applyNestedMainHeaderStyle(mainCell, config, field.getExcel());

                XSSFCell subCell = subHeaderRow.createCell(currentColumn);
                subCell.setCellValue("");
                applyNestedSubHeaderStyle(subCell);

                // 合并主表头和次级表头的单元格，让普通字段的表头跨两行显示
                CellRangeAddress mergeRegion = new CellRangeAddress(0, 1, currentColumn, currentColumn);
                sheet.addMergedRegion(mergeRegion);

                currentColumn++;
            }
        }

        // 应用表头样式到整行
        applyHeaderStyleToRow(sheet.getRow(0), config);
        applySubHeaderStyleToRow(sheet.getRow(1));
    }

    /**
     * 写入数据
     */
    private void writeData(XSSFSheet sheet, List<?> dataList) {
        int currentRow = context.getDataStartRow();

        for (Object data : dataList) {
            // 计算当前数据行需要的行数（基于最大List长度）
            int maxListSize = getMaxListSize(data);

            int currentColumn = 0;

            for (ExportField field : context.getExportFields()) {
                if (field.isMap()) {
                    // Map类型横向扩展
                    Map<?, ?> map = getMapValue(data, field);
                    if (map != null && !map.isEmpty()) {
                        List<?> allKeys = MapProcessor.getAllKeys(dataList, field);
                        for (Object key : allKeys) {
                            XSSFCell cell = ensureRowAndGetCell(sheet, currentRow, currentColumn);
                            Object value = map.get(key);
                            cell.setCellValue(value != null ? value.toString() : "");
                            applyCellStyle(cell, field.getExcel());
                            currentColumn++;
                        }
                    } else {
                        // Map为空或null，填充空单元格
                        for (int i = 0; i < getMaxMapColumns(dataList, field); i++) {
                            XSSFCell cell = ensureRowAndGetCell(sheet, currentRow, currentColumn);
                            cell.setCellValue("");
                            currentColumn++;
                        }
                    }
                } else if (field.isNested()) {
                    // 嵌套对象类型
                    NestedProcessor.NestedFieldInfo nestedInfo = getNestedFieldInfo(field);
                    if (nestedInfo != null) {
                        Object nestedObj = ReflectionUtil.getFieldValue(data, field.getField());
                        for (int i = 0; i < nestedInfo.childHeaders.size(); i++) {
                            XSSFCell cell = ensureRowAndGetCell(sheet, currentRow, currentColumn);
                            String value = "";
                            if (nestedObj != null && i < nestedInfo.childFields.size()) {
                                Object fieldValue = ReflectionUtil.getFieldValue(nestedObj, nestedInfo.childFields.get(i).getField());
                                value = fieldValue != null ? fieldValue.toString() : "";
                            }
                            cell.setCellValue(value);
                            applyCellStyle(cell, field.getExcel());
                            currentColumn++;
                        }
                    } else {
                        // 递归模式或其他情况
                        Object value = ReflectionUtil.getFieldValue(data, field.getField());
                        XSSFCell cell = ensureRowAndGetCell(sheet, currentRow, currentColumn);
                        cell.setCellValue(value != null ? value.toString() : "");
                        applyCellStyle(cell, field.getExcel());
                        currentColumn++;
                    }
                } else if (field.isList()) {
                    // List类型
                    List<?> list = getListValue(data, field);
                    ListFieldInfo listInfo = getListFieldInfo(field);

                    if (listInfo != null && listInfo.columnSpan > 0) {
                        // 需要横向展开次级列
                        int maxItems = list != null ? list.size() : 0;
                        for (int i = 0; i < listInfo.columnSpan; i++) {
                            // 获取子字段
                            ExportField childField = listInfo.childFields.get(i);
                            for (int rowOffset = 0; rowOffset < maxItems; rowOffset++) {
                                XSSFCell cell = ensureRowAndGetCell(sheet, currentRow + rowOffset, currentColumn + i);
                                String value = "";
                                if (rowOffset < list.size()) {
                                    Object item = list.get(rowOffset);
                                    if (item != null) {
                                        value = getListItemFieldValue(item, childField);
                                    }
                                }
                                cell.setCellValue(value);
                                applyCellStyle(cell, listInfo.parentField.getExcel());
                            }
                        }
                        currentColumn += listInfo.columnSpan;
                    } else {
                        // 纵向扩展（不展开列）
                        if (list != null && !list.isEmpty()) {
                            for (int i = 0; i < list.size(); i++) {
                                XSSFCell cell = ensureRowAndGetCell(sheet, currentRow + i, currentColumn);
                                Object item = list.get(i);
                                String value = getListItemValue(item);
                                cell.setCellValue(value);
                                applyCellStyle(cell, field.getExcel());
                            }
                        }
                        currentColumn++;
                    }
                } else {
                    // 普通字段（纵向合并）
                    String value = getFieldValue(data, field);
                    XSSFCell cell = ensureRowAndGetCell(sheet, currentRow, currentColumn);
                    cell.setCellValue(value);
                    applyCellStyle(cell, field.getExcel());
                    currentColumn++;
                }
            }

            currentRow += Math.max(maxListSize, 1);
        }
    }

    /**
     * 确保行存在并获取单元格
     */
    private XSSFCell ensureRowAndGetCell(XSSFSheet sheet, int rowIndex, int columnIndex) {
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
     * 获取最大List长度
     */
    private int getMaxListSize(Object data) {
        int maxSize = 1;
        for (ExportField field : context.getExportFields()) {
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
     * 获取最大Map列数
     */
    private int getMaxMapColumns(List<?> dataList, ExportField field) {
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
     * 获取Map值
     */
    private Map<?, ?> getMapValue(Object data, ExportField field) {
        try {
            Object value = ReflectionUtil.getFieldValue(data, field.getField());
            if (value instanceof Map) {
                return (Map<?, ?>) value;
            }
        } catch (Exception e) {
            // ignore
        }
        return null;
    }

    /**
     * 获取List值
     */
    private List<?> getListValue(Object data, ExportField field) {
        try {
            Object value = ReflectionUtil.getFieldValue(data, field.getField());
            if (value instanceof List) {
                return (List<?>) value;
            }
        } catch (Exception e) {
            // ignore
        }
        return null;
    }

    /**
     * 应用单元格合并
     */
    private void applyMerges(XSSFSheet sheet, List<?> dataList) {
        // 合并表头（Map类型和嵌套对象的主表头已在writeHeader中合并）

        // 合并数据单元格（纵向合并）
        int currentRow = context.getDataStartRow();

        for (Object data : dataList) {
            int maxListSize = getMaxListSize(data);
            int currentColumn = 0;

            for (ExportField field : context.getExportFields()) {
                if (field.isMap()) {
                    // Map类型横向扩展，每个key占一列，不需要纵向合并
                    List<Object> allKeys = MapProcessor.getAllKeys(dataList, field);
                    currentColumn += allKeys.size();
                } else if (field.isNested()) {
                    // 嵌套对象类型 - 对展开后的每个子列进行纵向合并
                    NestedProcessor.NestedFieldInfo nestedInfo = getNestedFieldInfo(field);
                    if (nestedInfo != null) {
                        for (int i = 0; i < nestedInfo.childHeaders.size(); i++) {
                            if (maxListSize > 1) {
                                CellRangeAddress mergeRegion = new CellRangeAddress(
                                        currentRow, currentRow + maxListSize - 1,
                                        currentColumn, currentColumn
                                );
                                sheet.addMergedRegion(mergeRegion);
                                // 为合并区域内所有单元格设置边框
                                applyBordersToMergedRegion(sheet, mergeRegion);
                            }
                            currentColumn++;
                        }
                    } else {
                        // 递归模式或其他情况，只有一列
                        if (maxListSize > 1) {
                            CellRangeAddress mergeRegion = new CellRangeAddress(
                                    currentRow, currentRow + maxListSize - 1,
                                    currentColumn, currentColumn
                            );
                            sheet.addMergedRegion(mergeRegion);
                            applyBordersToMergedRegion(sheet, mergeRegion);
                        }
                        currentColumn++;
                    }
                } else if (field.isList()) {
                    // List类型
                    ListFieldInfo listInfo = getListFieldInfo(field);
                    if (listInfo != null && listInfo.columnSpan > 0) {
                        // List展开的列，不需要纵向合并（每个Item占一行）
                        currentColumn += listInfo.columnSpan;
                    } else {
                        // 不展开的List字段，不需要纵向合并
                        currentColumn++;
                    }
                } else {
                    // 普通字段纵向合并
                    if (maxListSize > 1) {
                        CellRangeAddress mergeRegion = new CellRangeAddress(
                                currentRow, currentRow + maxListSize - 1,
                                currentColumn, currentColumn
                        );
                        sheet.addMergedRegion(mergeRegion);
                        applyBordersToMergedRegion(sheet, mergeRegion);
                    }
                    currentColumn++;
                }
            }

            currentRow += Math.max(maxListSize, 1);
        }
    }

    /**
     * 为合并区域内的所有单元格设置边框
     */
    private void applyBordersToMergedRegion(XSSFSheet sheet, CellRangeAddress region) {
        XSSFWorkbook workbook = sheet.getWorkbook();

        // 遍历合并区域内的所有单元格
        for (int rowIdx = region.getFirstRow(); rowIdx <= region.getLastRow(); rowIdx++) {
            for (int colIdx = region.getFirstColumn(); colIdx <= region.getLastColumn(); colIdx++) {
                XSSFRow row = sheet.getRow(rowIdx);
                if (row == null) {
                    row = sheet.createRow(rowIdx);
                }
                XSSFCell cell = row.getCell(colIdx);
                if (cell == null) {
                    cell = row.createCell(colIdx);
                }

                // 为单元格添加边框
                XSSFCellStyle borderStyle = workbook.createCellStyle();
                // 复制原有样式
                if (cell.getCellStyle() != null) {
                    borderStyle.cloneStyleFrom(cell.getCellStyle());
                }
                // 设置边框
                borderStyle.setBorderTop(BorderStyle.THIN);
                borderStyle.setBorderBottom(BorderStyle.THIN);
                borderStyle.setBorderLeft(BorderStyle.THIN);
                borderStyle.setBorderRight(BorderStyle.THIN);
                cell.setCellStyle(borderStyle);
            }
        }
    }

    /**
     * 为表头合并区域内的所有单元格设置边框
     */
    private void applyHeaderBordersToMergedRegion(XSSFSheet sheet, CellRangeAddress region) {
        XSSFWorkbook workbook = sheet.getWorkbook();

        // 遍历合并区域内的所有单元格
        for (int rowIdx = region.getFirstRow(); rowIdx <= region.getLastRow(); rowIdx++) {
            for (int colIdx = region.getFirstColumn(); colIdx <= region.getLastColumn(); colIdx++) {
                XSSFRow row = sheet.getRow(rowIdx);
                if (row == null) {
                    row = sheet.createRow(rowIdx);
                }
                XSSFCell cell = row.getCell(colIdx);
                if (cell == null) {
                    cell = row.createCell(colIdx);
                }

                // 为单元格添加边框
                XSSFCellStyle borderStyle = workbook.createCellStyle();
                // 复制原有样式
                if (cell.getCellStyle() != null) {
                    borderStyle.cloneStyleFrom(cell.getCellStyle());
                }
                // 设置边框
                borderStyle.setBorderTop(BorderStyle.THIN);
                borderStyle.setBorderBottom(BorderStyle.THIN);
                borderStyle.setBorderLeft(BorderStyle.THIN);
                borderStyle.setBorderRight(BorderStyle.THIN);
                cell.setCellStyle(borderStyle);
            }
        }
    }

    /**
     * 设置列宽
     */
    private void applyColumnWidths(XSSFSheet sheet) {
        // 自动调整所有列宽
        int maxColumns = sheet.getRow(0) != null ? sheet.getRow(0).getLastCellNum() : 0;
        for (int i = 0; i < maxColumns; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * 获取表头文本
     */
    private String getHeaderText(ExportField field) {
        if (field.getExcel() != null && !field.getExcel().header().isEmpty()) {
            return field.getExcel().header();
        }
        return field.getField().getName();
    }

    /**
     * 获取字段值
     */
    private String getFieldValue(Object data, ExportField field) {
        try {
            Object value = ReflectionUtil.getFieldValue(data, field.getField());
            if (value == null) {
                return "";
            }

            Excel excel = field.getExcel();
            if (excel != null) {
                if (value instanceof Date) {
                    return ReflectionUtil.formatDate((Date) value, excel.dateFormat());
                }
                if (value instanceof LocalDate) {
                    return ReflectionUtil.formatLocalDate((LocalDate) value, excel.dateFormat());
                }
                if (value instanceof LocalDateTime) {
                    return ReflectionUtil.formatLocalDateTime((LocalDateTime) value, excel.dateFormat());
                }
                if (value instanceof Number && !excel.numberFormat().isEmpty()) {
                    return ReflectionUtil.formatNumber(value, excel.numberFormat());
                }
            }

            return value.toString();
        } catch (Exception e) {
            return "Error: " + e.getMessage();
        }
    }

    /**
     * 获取List元素的值
     * 如果元素是简单类型，直接返回toString()
     * 如果元素是复杂类型，提取其@Excel注解的属性值
     */
    private String getListItemValue(Object item) {
        if (item == null) {
            return "";
        }

        // 尝试获取泛型类型中标注了@Excel注解的字段
        Class<?> itemType = item.getClass();
        java.lang.reflect.Field[] fields = itemType.getDeclaredFields();

        // 收集所有@Excel注解的字段值
        List<String> values = new ArrayList<>();
        for (java.lang.reflect.Field f : fields) {
            Excel excel = f.getAnnotation(Excel.class);
            if (excel != null) {
                try {
                    f.setAccessible(true);
                    Object value = f.get(item);
                    if (value != null) {
                        if (value instanceof Date) {
                            values.add(ReflectionUtil.formatDate((Date) value, excel.dateFormat()));
                        } else if (value instanceof Number && !excel.numberFormat().isEmpty()) {
                            values.add(ReflectionUtil.formatNumber(value, excel.numberFormat()));
                        } else {
                            values.add(value.toString());
                        }
                    } else {
                        values.add("");
                    }
                } catch (Exception e) {
                    values.add("");
                }
            }
        }

        // 如果没有找到@Excel注解的字段，返回toString()
        if (values.isEmpty()) {
            return item.toString();
        }

        // 用分隔符拼接各字段值
        return String.join(", ", values);
    }

    /**
     * 获取List元素的指定字段值
     */
    private String getListItemFieldValue(Object item, ExportField field) {
        if (item == null) {
            return "";
        }
        try {
            java.lang.reflect.Field f = field.getField();
            f.setAccessible(true);
            Object value = f.get(item);
            if (value == null) {
                return "";
            }
            Excel excel = field.getExcel();
            if (excel != null) {
                if (value instanceof Date) {
                    return ReflectionUtil.formatDate((Date) value, excel.dateFormat());
                }
                if (value instanceof Number && !excel.numberFormat().isEmpty()) {
                    return ReflectionUtil.formatNumber(value, excel.numberFormat());
                }
            }
            return value.toString();
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * 应用表头样式到单元格
     */
    private void applyHeaderStyle(XSSFCell cell) {
        XSSFWorkbook workbook = cell.getSheet().getWorkbook();
        XSSFCellStyle style = workbook.createCellStyle();
        
        // XSSF 使用 XSSFColor 对象设置颜色
        XSSFColor bgColor = parseXssfColor("#4472C4");
        if (bgColor != null) {
            style.setFillForegroundColor(bgColor);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        XSSFFont font = workbook.createFont();
        XSSFColor fontColor = parseXssfColor("#FFFFFF");
        if (fontColor != null) {
            font.setColor(fontColor);
        }
        font.setBold(true);
        style.setFont(font);

        // 边框
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        cell.setCellStyle(style);
    }

    /**
     * 应用嵌套主表头样式
     */
    private void applyNestedMainHeaderStyle(XSSFCell cell, ExcelConfig config, Excel excel) {
        XSSFWorkbook workbook = cell.getSheet().getWorkbook();
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        // 背景色
        String bgColor = "#4472C4";
        if (excel != null && !excel.nestedHeaderBgColor().isEmpty()) {
            bgColor = excel.nestedHeaderBgColor();
        } else if (config != null && !config.headerBackgroundColor().isEmpty()) {
            bgColor = config.headerBackgroundColor();
        }
        XSSFColor bgXssfColor = parseXssfColor(bgColor);
        if (bgXssfColor != null) {
            style.setFillForegroundColor(bgXssfColor);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        // 字体
        XSSFFont font = workbook.createFont();
        String fontColor = "#FFFFFF";
        if (config != null && !config.headerFontColor().isEmpty()) {
            fontColor = config.headerFontColor();
        }
        XSSFColor fontXssfColor = parseXssfColor(fontColor);
        if (fontXssfColor != null) {
            font.setColor(fontXssfColor);
        }
        font.setBold(true);
        style.setFont(font);

        // 边框
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        cell.setCellStyle(style);
    }

    /**
     * 应用嵌套次级表头样式
     */
    private void applyNestedSubHeaderStyle(XSSFCell cell) {
        XSSFWorkbook workbook = cell.getSheet().getWorkbook();
        XSSFCellStyle style = workbook.createCellStyle();
        
        XSSFColor bgColor = parseXssfColor("#D6DCE5");
        if (bgColor != null) {
            style.setFillForegroundColor(bgColor);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        XSSFFont font = workbook.createFont();
        XSSFColor fontColor = parseXssfColor("#000000");
        if (fontColor != null) {
            font.setColor(fontColor);
        }
        font.setBold(true);
        style.setFont(font);

        // 边框
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        cell.setCellStyle(style);
    }

    /**
     * 应用表头样式到整行
     */
    private void applyHeaderStyleToRow(XSSFRow row, ExcelConfig config) {
        if (row == null) return;

        String bgColor = "#4472C4";
        String fontColor = "#FFFFFF";
        boolean bold = true;

        if (config != null) {
            if (!config.headerBackgroundColor().isEmpty()) {
                bgColor = config.headerBackgroundColor();
            }
            if (!config.headerFontColor().isEmpty()) {
                fontColor = config.headerFontColor();
            }
            bold = config.headerFontBold();
        }

        XSSFWorkbook workbook = row.getSheet().getWorkbook();
        XSSFCellStyle style = workbook.createCellStyle();
        
        XSSFColor bgXssfColor = parseXssfColor(bgColor);
        if (bgXssfColor != null) {
            style.setFillForegroundColor(bgXssfColor);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        XSSFFont font = workbook.createFont();
        XSSFColor fontXssfColor = parseXssfColor(fontColor);
        if (fontXssfColor != null) {
            font.setColor(fontXssfColor);
        }
        font.setBold(bold);
        style.setFont(font);

        // 边框
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        for (int i = 0; i < row.getLastCellNum(); i++) {
            XSSFCell cell = row.getCell(i);
            if (cell != null) {
                cell.setCellStyle(style);
            }
        }
    }

    /**
     * 应用次级表头样式到整行
     */
    private void applySubHeaderStyleToRow(XSSFRow row) {
        if (row == null) return;

        for (int i = 0; i < row.getLastCellNum(); i++) {
            XSSFCell cell = row.getCell(i);
            if (cell != null && cell.getCellStyle() == null) {
                applyNestedSubHeaderStyle(cell);
            }
        }
    }

    /**
     * 应用单元格样式
     */
    private void applyCellStyle(XSSFCell cell, Excel excel) {
        XSSFWorkbook workbook = cell.getSheet().getWorkbook();
        XSSFCellStyle style = workbook.createCellStyle();

        // 字体颜色和粗体（默认黑色字体）
        XSSFFont font = workbook.createFont();
        font.setColor(IndexedColors.BLACK.getIndex());
        font.setBold(false);

        if (excel != null) {
            // 背景色
            if (!excel.backgroundColor().isEmpty()) {
                try {
                    XSSFColor color = parseXssfColor(excel.backgroundColor());
                    if (color != null) {
                        style.setFillForegroundColor(color);
                        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    }
                } catch (Exception e) {
                    // 忽略颜色设置错误
                }
            }

            // 字体颜色
            if (!excel.fontColor().isEmpty()) {
                try {
                    XSSFColor color = parseXssfColor(excel.fontColor());
                    if (color != null) {
                        font.setColor(color);
                    }
                } catch (Exception e) {
                    // 忽略颜色设置错误
                }
            }
            font.setBold(excel.fontBold());

            // 对齐方式
            style.setAlignment(convertHorizontalAlign(excel.horizontalAlign()));
            style.setVerticalAlignment(convertVerticalAlign(excel.verticalAlign()));
        }

        style.setFont(font);

        // 边框
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        cell.setCellStyle(style);
    }

    /**
     * 将十六进制颜色字符串转换为 XSSFColor
     */
    private XSSFColor parseXssfColor(String hex) {
        if (hex == null || hex.isEmpty()) {
            return null;
        }
        try {
            String color = hex.replace("#", "");
            if (color.length() == 6) {
                int r = Integer.parseInt(color.substring(0, 2), 16);
                int g = Integer.parseInt(color.substring(2, 4), 16);
                int b = Integer.parseInt(color.substring(4, 6), 16);
                return new XSSFColor(new java.awt.Color(r, g, b), null);
            }
        } catch (Exception e) {
            // 忽略
        }
        return null;
    }

    /**
     * 转换水平对齐方式
     */
    private HorizontalAlignment convertHorizontalAlign(Excel.HorizontalAlign align) {
        switch (align) {
            case LEFT: return HorizontalAlignment.LEFT;
            case CENTER: return HorizontalAlignment.CENTER;
            case RIGHT: return HorizontalAlignment.RIGHT;
            case JUSTIFIED: return HorizontalAlignment.JUSTIFY;
            default: return HorizontalAlignment.CENTER;
        }
    }

    /**
     * 转换垂直对齐方式
     */
    private VerticalAlignment convertVerticalAlign(Excel.VerticalAlign align) {
        switch (align) {
            case TOP: return VerticalAlignment.TOP;
            case CENTER: return VerticalAlignment.CENTER;
            case BOTTOM: return VerticalAlignment.BOTTOM;
            default: return VerticalAlignment.CENTER;
        }
    }
}
