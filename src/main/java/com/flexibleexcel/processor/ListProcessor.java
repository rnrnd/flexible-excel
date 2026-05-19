package com.flexibleexcel.processor;

import com.flexibleexcel.annotation.Excel;
import com.flexibleexcel.core.CellData;
import com.flexibleexcel.core.ExcelContext;
import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.util.ReflectionUtil;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * List类型字段处理器
 * 处理List类型的属性，支持纵向扩展
 */
public class ListProcessor {

    private final ExcelContext context;

    public ListProcessor(ExcelContext context) {
        this.context = context;
    }

    /**
     * 处理List字段，获取展开后的单元格数据
     *
     * @param data   包含List的对象
     * @param field  List类型的字段
     * @param excel  字段上的Excel注解
     * @param row    数据所在的行索引
     * @param column 字段所在的列索引
     * @return 展开后的单元格数据列表
     */
    public List<CellData> process(Object data, ExportField field, int row, int column) {
        List<CellData> cellDataList = new ArrayList<>();

        try {
            Object value = ReflectionUtil.getFieldValue(data, field.getField());
            if (value == null) {
                // List为null时，创建一个空单元格（跨行合并）
                cellDataList.add(CellData.merged("", row, row + context.getMaxRows() - 1, column, column));
                return cellDataList;
            }

            List<?> list = (List<?>) value;
            if (list.isEmpty()) {
                // List为空时，也创建空单元格
                cellDataList.add(CellData.merged("", row, row + context.getMaxRows() - 1, column, column));
                return cellDataList;
            }

            // 为List的每个元素创建一个单元格（纵向扩展）
            for (int i = 0; i < list.size(); i++) {
                Object item = list.get(i);
                String cellValue = ReflectionUtil.toString(item, getDateFormat(field));
                CellData cellData = CellData.of(cellValue, row + i, column);
                cellDataList.add(cellData);
            }

        } catch (Exception e) {
            cellDataList.add(CellData.of("Error: " + e.getMessage(), row, column));
        }

        return cellDataList;
    }

    /**
     * 获取当前列中List的最大长度
     */
    public static int getMaxListSize(List<?> dataList, ExportField field) {
        int maxSize = 1; // 最小为1
        for (Object data : dataList) {
            try {
                Object value = ReflectionUtil.getFieldValue(data, field.getField());
                if (value instanceof List) {
                    List<?> list = (List<?>) value;
                    maxSize = Math.max(maxSize, list.size());
                }
            } catch (Exception e) {
                // ignore
            }
        }
        return maxSize;
    }

    /**
     * 计算需要合并的行数
     */
    public static int calculateMergeRows(int listSize, int maxListSize) {
        if (listSize == 0) {
            return maxListSize;
        }
        return maxListSize - listSize;
    }

    /**
     * 获取日期格式化字符串
     */
    private String getDateFormat(ExportField field) {
        if (field.getExcel() != null) {
            return field.getExcel().dateFormat();
        }
        return "yyyy-MM-dd";
    }
}
