package com.flexibleexcel.core.structure;

import com.flexibleexcel.core.ExcelContext;
import com.flexibleexcel.util.ReflectionUtil;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Map 结构辅助工具。
 */
public final class MapKeyCollector {

    private MapKeyCollector() {
    }

    /**
     * 获取数据集中某个 Map 字段的所有 key，并按可比较顺序排序。
     */
    public static List<Object> getAllKeys(List<?> dataList, ExcelContext.ExportField field) {
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
        try {
            if (keys.stream().allMatch(k -> k instanceof Comparable)) {
                @SuppressWarnings("unchecked")
                List<Comparable<Object>> comparableKeys = (List<Comparable<Object>>) (List<?>) keys;
                Collections.sort(comparableKeys);
            }
        } catch (Exception e) {
            // keep original order
        }
        return keys;
    }
}
