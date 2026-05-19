package com.flexibleexcel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel导出字段注解
 * 用于标记类的属性为可导出字段
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {

    /**
     * 主表头名称
     */
    String header() default "";

    /**
     * 列宽（单位：字符数）
     * 默认值为-1，表示使用自动宽度
     */
    int columnWidth() default -1;

    /**
     * 行高（单位：点）
     * 默认值为-1，表示使用默认行高
     */
    int rowHeight() default -1;

    /**
     * 背景色（RGB格式，如 "#FF0000" 表示红色）
     */
    String backgroundColor() default "";

    /**
     * 字体颜色（RGB格式）
     */
    String fontColor() default "";

    /**
     * 是否加粗
     */
    boolean fontBold() default false;

    /**
     * 水平对齐方式
     * LEFT, CENTER, RIGHT
     */
    HorizontalAlign horizontalAlign() default HorizontalAlign.CENTER;

    /**
     * 垂直对齐方式
     * TOP, CENTER, BOTTOM
     */
    VerticalAlign verticalAlign() default VerticalAlign.CENTER;

    /**
     * 日期格式化（当字段为Date类型时有效）
     * 格式如 "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss"
     */
    String dateFormat() default "yyyy-MM-dd";

    /**
     * 数字格式化（当字段为数字类型时有效）
     * 格式如 "#,##0.00"
     */
    String numberFormat() default "";

    /**
     * 导出顺序（从小到大排序）
     */
    int order() default 0;

    /**
     * 是否为次级表头（用于Map类型展开后的列）
     */
    boolean subHeader() default false;

    /**
     * 嵌套展开模式（用于嵌套对象类型）
     * NONE: 不展开，直接toString()
     * HORIZONTAL: 横向展开，子属性作为次级表头横向排列
     * RECURSIVE: 递归展开，深度展开所有嵌套的@Excel属性
     */
    NestedMode nested() default NestedMode.NONE;

    /**
     * 嵌套展开时的主表头背景色
     */
    String nestedHeaderBgColor() default "";

    /**
     * 嵌套展开时的次级表头背景色
     */
    String nestedSubHeaderBgColor() default "";

    /**
     * 嵌套展开时是否合并主表头单元格
     * 默认为true，横向合并主表头
     */
    boolean mergeNestedHeader() default true;

    /**
     * 水平对齐方式枚举
     */
    enum HorizontalAlign {
        LEFT,
        CENTER,
        RIGHT,
        JUSTIFIED
    }

    /**
     * 垂直对齐方式枚举
     */
    enum VerticalAlign {
        TOP,
        CENTER,
        BOTTOM
    }

    /**
     * 嵌套展开模式枚举
     */
    enum NestedMode {
        /**
         * 不展开，直接调用toString()
         */
        NONE,
        /**
         * 横向展开，子属性作为次级表头横向排列
         */
        HORIZONTAL,
        /**
         * 递归展开，深度展开所有嵌套的@Excel属性
         */
        RECURSIVE
    }
}
