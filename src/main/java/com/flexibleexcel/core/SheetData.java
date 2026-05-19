package com.flexibleexcel.core;

/**
 * Sheet数据封装类
 * 用于多Sheet导出时封装每个Sheet的数据
 */
public class SheetData {

    private String sheetName;
    private Class<?> clazz;
    private Object dataList;

    private SheetData(String sheetName, Class<?> clazz, Object dataList) {
        this.sheetName = sheetName;
        this.clazz = clazz;
        this.dataList = dataList;
    }

    /**
     * 创建SheetData
     *
     * @param sheetName Sheet名称
     * @param clazz     数据类（带@Excel注解）
     * @param dataList  数据列表
     * @return SheetData实例
     */
    public static SheetData of(String sheetName, Class<?> clazz, java.util.List<?> dataList) {
        return new SheetData(sheetName, clazz, dataList);
    }

    /**
     * 创建SheetData（使用类上的@ExcelConfig注解的sheetName）
     *
     * @param clazz    数据类（带@Excel注解）
     * @param dataList 数据列表
     * @return SheetData实例
     */
    public static SheetData of(Class<?> clazz, java.util.List<?> dataList) {
        com.flexibleexcel.annotation.ExcelConfig config = clazz.getAnnotation(com.flexibleexcel.annotation.ExcelConfig.class);
        String sheetName = (config != null && !config.sheetName().isEmpty()) 
                ? config.sheetName() 
                : clazz.getSimpleName();
        return new SheetData(sheetName, clazz, dataList);
    }

    public String getSheetName() {
        return sheetName;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    @SuppressWarnings("unchecked")
    public java.util.List<?> getDataList() {
        return (java.util.List<?>) dataList;
    }
}
