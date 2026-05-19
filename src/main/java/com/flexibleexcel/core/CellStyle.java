package com.flexibleexcel.core;

import com.flexibleexcel.annotation.Excel.HorizontalAlign;
import com.flexibleexcel.annotation.Excel.VerticalAlign;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 单元格样式配置
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class CellStyle {

    /**
     * 背景色（RGB格式）
     */
    private String backgroundColor;

    /**
     * 字体颜色（RGB格式）
     */
    private String fontColor;

    /**
     * 是否加粗
     */
    private boolean fontBold;

    /**
     * 水平对齐
     */
    private HorizontalAlign horizontalAlign;

    /**
     * 垂直对齐
     */
    private VerticalAlign verticalAlign;

    /**
     * 列宽
     */
    private int columnWidth;

    /**
     * 行高
     */
    private int rowHeight;

    /**
     * 日期格式化
     */
    private String dateFormat;

    /**
     * 数字格式化
     */
    private String numberFormat;

    /**
     * 是否设置边框
     */
    private boolean border;

    /**
     * 边框颜色
     */
    private String borderColor;

    /**
     * 字体大小
     */
    private int fontSize;

    /**
     * 创建默认样式
     */
    public static CellStyle defaultStyle() {
        return CellStyle.builder()
                .horizontalAlign(HorizontalAlign.CENTER)
                .verticalAlign(VerticalAlign.CENTER)
                .fontBold(false)
                .border(true)
                .borderColor("#000000")
                .fontSize(11)
                .build();
    }
}
