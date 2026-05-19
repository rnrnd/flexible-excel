package com.flexibleexcel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel导出配置注解
 * 用于配置整个导出行为
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelConfig {

    /**
     * 工作表名称
     */
    String sheetName() default "Sheet1";

    /**
     * 表头行数（默认1行）
     */
    int headerRows() default 1;

    /**
     * 是否自动生成表头
     */
    boolean autoGenerateHeader() default true;

    /**
     * 表头背景色
     */
    String headerBackgroundColor() default "#4472C4";

    /**
     * 表头字体颜色
     */
    String headerFontColor() default "#FFFFFF";

    /**
     * 表头是否加粗
     */
    boolean headerFontBold() default true;

    /**
     * 数据起始行（0-based）
     */
    int dataStartRow() default 1;

    /**
     * 数据起始列（0-based）
     */
    int dataStartColumn() default 0;

    /**
     * 是否启用交替行颜色
     */
    boolean alternateRowColors() default false;

    /**
     * 奇数行背景色
     */
    String oddRowColor() default "";

    /**
     * 偶数行背景色
     */
    String evenRowColor() default "";

    /**
     * 是否导出时间戳
     */
    boolean exportTimestamp() default false;
}
