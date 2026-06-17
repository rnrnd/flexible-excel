package com.flexibleexcel.core;

import com.flexibleexcel.annotation.Excel;
import com.flexibleexcel.annotation.ExcelConfig;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 样式辅助类，负责表头、数据和合并单元格边框样式处理。
 */
class StyleHelper {

    /**
     * 应用表头样式到单元格。
     */
    void applyHeaderStyle(XSSFCell cell) {
        XSSFWorkbook workbook = cell.getSheet().getWorkbook();
        XSSFCellStyle style = workbook.createCellStyle();

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

        applyThinBorders(style);
        cell.setCellStyle(style);
    }

    /**
     * 应用嵌套主表头样式。
     */
    void applyNestedMainHeaderStyle(XSSFCell cell, ExcelConfig config, Excel excel) {
        XSSFWorkbook workbook = cell.getSheet().getWorkbook();
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

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

        applyThinBorders(style);
        cell.setCellStyle(style);
    }

    /**
     * 应用嵌套次级表头样式。
     */
    void applyNestedSubHeaderStyle(XSSFCell cell) {
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

        applyThinBorders(style);
        cell.setCellStyle(style);
    }

    /**
     * 应用表头样式到整行。
     */
    void applyHeaderStyleToRow(XSSFRow row, ExcelConfig config) {
        if (row == null) {
            return;
        }

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

        applyThinBorders(style);

        for (int i = 0; i < row.getLastCellNum(); i++) {
            XSSFCell cell = row.getCell(i);
            if (cell != null) {
                cell.setCellStyle(style);
            }
        }
    }

    /**
     * 应用次级表头样式到整行。
     */
    void applySubHeaderStyleToRow(XSSFRow row) {
        if (row == null) {
            return;
        }

        for (int i = 0; i < row.getLastCellNum(); i++) {
            XSSFCell cell = row.getCell(i);
            if (cell != null && cell.getCellStyle() == null) {
                applyNestedSubHeaderStyle(cell);
            }
        }
    }

    /**
     * 应用单元格样式。
     */
    void applyCellStyle(XSSFCell cell, Excel excel) {
        XSSFWorkbook workbook = cell.getSheet().getWorkbook();
        XSSFCellStyle style = workbook.createCellStyle();

        XSSFFont font = workbook.createFont();
        font.setColor(IndexedColors.BLACK.getIndex());
        font.setBold(false);

        if (excel != null) {
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
            style.setAlignment(convertHorizontalAlign(excel.horizontalAlign()));
            style.setVerticalAlignment(convertVerticalAlign(excel.verticalAlign()));
        }

        style.setFont(font);
        applyThinBorders(style);
        cell.setCellStyle(style);
    }

    /**
     * 为合并区域内的所有单元格设置边框。
     */
    void applyBordersToMergedRegion(XSSFSheet sheet, CellRangeAddress region) {
        applyBordersToMergedRegion(sheet, region, true);
    }

    /**
     * 为表头合并区域内的所有单元格设置边框。
     */
    void applyHeaderBordersToMergedRegion(XSSFSheet sheet, CellRangeAddress region) {
        applyBordersToMergedRegion(sheet, region, true);
    }

    /**
     * 为合并区域内的单元格设置边框。
     */
    void applyBordersToMergedRegion(XSSFSheet sheet, CellRangeAddress region, boolean preserveStyle) {
        XSSFWorkbook workbook = sheet.getWorkbook();

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

                XSSFCellStyle borderStyle = workbook.createCellStyle();
                if (preserveStyle && cell.getCellStyle() != null) {
                    borderStyle.cloneStyleFrom(cell.getCellStyle());
                }
                applyThinBorders(borderStyle);
                cell.setCellStyle(borderStyle);
            }
        }
    }

    /**
     * 为样式添加细边框。
     */
    private void applyThinBorders(XSSFCellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    /**
     * 将十六进制颜色字符串转换为 XSSFColor。
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
     * 转换水平对齐方式。
     */
    private HorizontalAlignment convertHorizontalAlign(Excel.HorizontalAlign align) {
        switch (align) {
            case LEFT:
                return HorizontalAlignment.LEFT;
            case CENTER:
                return HorizontalAlignment.CENTER;
            case RIGHT:
                return HorizontalAlignment.RIGHT;
            case JUSTIFIED:
                return HorizontalAlignment.JUSTIFY;
            default:
                return HorizontalAlignment.CENTER;
        }
    }

    /**
     * 转换垂直对齐方式。
     */
    private VerticalAlignment convertVerticalAlign(Excel.VerticalAlign align) {
        switch (align) {
            case TOP:
                return VerticalAlignment.TOP;
            case CENTER:
                return VerticalAlignment.CENTER;
            case BOTTOM:
                return VerticalAlignment.BOTTOM;
            default:
                return VerticalAlignment.CENTER;
        }
    }
}
