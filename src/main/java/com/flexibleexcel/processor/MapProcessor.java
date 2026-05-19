package com.flexibleexcel.processor;

import com.flexibleexcel.annotation.Excel;
import com.flexibleexcel.core.CellData;
import com.flexibleexcel.core.ExcelContext;
import com.flexibleexcel.core.ExcelContext.ExportField;
import com.flexibleexcel.util.ReflectionUtil;

import java.util.*;

/**
 * Map类型字段处理器
 * 处理Map类型的属性，支持横向扩展
 */
public class MapProcessor {

    private final ExcelContext context;

    public MapProcessor(ExcelContext context) {
        this.context = context;
    }

    /**
     * 处理Map字段，获取展开后的单元格数据
     *
     * @param data   包含Map的对象
     * @param field  Map类型的字段
     * @param row    数据所在的行索引
     * @param column 字段所在的列索引（主表头列）
     * @return MapEntry列表，每个Entry包含key和value的CellData
     */
    public List<MapEntry> process(Object data, ExportField field, int row, int column) {
        List<MapEntry> entries = new ArrayList<>();

        try {
            Object value = ReflectionUtil.getFieldValue(data, field.getField());
            if (value == null) {
                // Map为null时，返回空列表
                return entries;
            }

            Map<?, ?> map = (Map<?, ?>) value;
            if (map.isEmpty()) {
                return entries;
            }

            // 获取排序后的key列表（确保顺序一致）
            List<?> sortedKeys = getSortedKeys(map);

            // 为每个key创建单元格数据
            int currentColumn = column;
            for (Object key : sortedKeys) {
                Object mapValue = map.get(key);
                String keyText = key != null ? key.toString() : "";
                String valueText = mapValue != null ? ReflectionUtil.toString(mapValue, getDateFormat(field)) : "";

                // 创建次级表头单元格（key）
                CellData headerCell = CellData.of(keyText, row, currentColumn);

                // 创建数据单元格（value）
                CellData valueCell = CellData.of(valueText, row, currentColumn);

                entries.add(new MapEntry(keyText, headerCell, valueCell));
                currentColumn++;
            }

        } catch (Exception e) {
            CellData errorCell = CellData.of("Error: " + e.getMessage(), row, column);
            entries.add(new MapEntry("error", errorCell, errorCell));
        }

        return entries;
    }

    /**
     * 获取排序后的Map key列表
     */
    private List<?> getSortedKeys(Map<?, ?> map) {
        List<Object> keys = new ArrayList<>(map.keySet());
        try {
            // 尝试自然排序
            if (keys.stream().allMatch(k -> k instanceof Comparable)) {
                @SuppressWarnings("unchecked")
                List<Comparable<Object>> comparableKeys = (List<Comparable<Object>>) (List<?>) keys;
                Collections.sort(comparableKeys);
            }
        } catch (Exception e) {
            // 如果排序失败，保持原始顺序
        }
        return keys;
    }

    /**
     * 获取Map展开后的列数
     */
    public static int getMapColumnCount(List<?> dataList, ExportField field) {
        int maxColumns = 1;
        for (Object data : dataList) {
            try {
                Object value = ReflectionUtil.getFieldValue(data, field.getField());
                if (value instanceof Map) {
                    Map<?, ?> map = (Map<?, ?>) value;
                    maxColumns = Math.max(maxColumns, map.size());
                }
            } catch (Exception e) {
                // ignore
            }
        }
        return maxColumns;
    }

    /**
     * 获取所有Map key的并集（用于表头）
     */
    public static List<Object> getAllKeys(List<?> dataList, ExportField field) {
        Set<Object> keySet = new LinkedHashSet<>();
        for (Object data : dataList) {
            try {
                Object value = ReflectionUtil.getFieldValue(data, field.getField());
                if (value instanceof Map) {
                    Map<?, ?> map = (Map<?, ?>) value;
                    keySet.addAll(map.keySet());
                }
            } catch (Exception e) {
                // ignore
            }
        }

        List<Object> keys = new ArrayList<>(keySet);
        // 尝试排序
        try {
            if (keys.stream().allMatch(k -> k instanceof Comparable)) {
                @SuppressWarnings("unchecked")
                List<Comparable<Object>> comparableKeys = (List<Comparable<Object>>) (List<?>) keys;
                Collections.sort(comparableKeys);
            }
        } catch (Exception e) {
            // 如果排序失败，保持原始顺序
        }
        return keys;
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

    /**
     * Map条目
     */
    public static class MapEntry {
        private final Object key;
        private final CellData headerCell;
        private final CellData valueCell;

        public MapEntry(Object key, CellData headerCell, CellData valueCell) {
            this.key = key;
            this.headerCell = headerCell;
            this.valueCell = valueCell;
        }

        public Object getKey() {
            return key;
        }

        public CellData getHeaderCell() {
            return headerCell;
        }

        public CellData getValueCell() {
            return valueCell;
        }
    }
}
