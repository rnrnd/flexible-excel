package com.flexibleexcel.core;

import com.flexibleexcel.annotation.ExcelConfig;
import com.flexibleexcel.core.structure.ListFieldInfo;
import com.flexibleexcel.core.structure.NestedProcessor;
import com.flexibleexcel.core.structure.StructureAnalyzer;
import com.flexibleexcel.core.write.DataWriter;
import com.flexibleexcel.core.write.HeaderWriter;
import com.flexibleexcel.core.write.MergeHandler;
import com.flexibleexcel.exception.ExcelExportException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel 导出器。
 * 负责组织导出流程，并协调表头写入、数据写入、合并策略与样式处理。
 */
public class ExcelExporter {
    private final StructureAnalyzer structureAnalyzer = new StructureAnalyzer();
    private final ExportSupport exportSupport = new ExportSupport();

    /**
     * 多 Sheet 导出时缓存的 Sheet 数据列表。
     */
    private final List<SheetData> sheetDataList = new ArrayList<>();

    /**
     * 创建导出器实例。
     */
    public static ExcelExporter create() {
        return new ExcelExporter();
    }

    /**
     * 添加一个 Sheet。
     *
     * @param sheetName Sheet 名称
     * @param clazz 数据模型类型
     * @param dataList 导出数据
     * @return 当前导出器实例
     */
    public ExcelExporter sheet(String sheetName, Class<?> clazz, List<?> dataList) {
        this.sheetDataList.add(SheetData.of(sheetName, clazz, dataList));
        return this;
    }

    /**
     * 添加一个 Sheet，名称优先取类上的 {@link ExcelConfig}。
     *
     * @param clazz 数据模型类型
     * @param dataList 导出数据
     * @return 当前导出器实例
     */
    public ExcelExporter sheet(Class<?> clazz, List<?> dataList) {
        this.sheetDataList.add(SheetData.of(clazz, dataList));
        return this;
    }

    /**
     * 导出多个 Sheet 到文件路径。
     *
     * @param filePath 输出文件路径
     */
    public void toFile(String filePath) {
        toFile(new File(filePath));
    }

    /**
     * 导出多个 Sheet 到文件。
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
     * 导出多个 Sheet 到输出流。
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
                renderSingleSheet(workbook, sheetData.getSheetName(), sheetData.getClazz(), sheetData.getDataList());
            }
            workbook.write(out);
            workbook.close();
        } catch (Exception e) {
            throw new ExcelExportException("导出Excel失败: " + e.getMessage(), e);
        }
    }

    /**
     * 导出单个 Sheet 到文件路径。
     *
     * @param clazz 数据模型类型
     * @param dataList 导出数据
     * @param filePath 输出文件路径
     */
    public void export(Class<?> clazz, List<?> dataList, String filePath) {
        export(clazz, dataList, new File(filePath));
    }

    /**
     * 导出单个 Sheet 到文件。
     *
     * @param clazz 数据模型类型
     * @param dataList 导出数据
     * @param file 输出文件
     */
    public void export(Class<?> clazz, List<?> dataList, File file) {
        try (OutputStream out = Files.newOutputStream(file.toPath())) {
            export(clazz, dataList, out);
        } catch (Exception e) {
            throw new ExcelExportException("导出Excel失败: " + e.getMessage(), e);
        }
    }

    /**
     * 导出单个 Sheet 到输出流。
     *
     * @param clazz 数据模型类型
     * @param dataList 导出数据
     * @param out 输出流
     */
    public void export(Class<?> clazz, List<?> dataList, OutputStream out) {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            renderSingleSheet(workbook, resolveSheetName(clazz), clazz, dataList);
            workbook.write(out);
            workbook.close();
        } catch (Exception e) {
            throw new ExcelExportException("导出Excel失败: " + e.getMessage(), e);
        }
    }

    /**
     * 渲染单个 Sheet。
     *
     * @param workbook 工作簿
     * @param sheetName Sheet 名称
     * @param clazz 数据模型类型
     * @param dataList 当前 Sheet 的数据
     */
    private void renderSingleSheet(XSSFWorkbook workbook, String sheetName, Class<?> clazz, List<?> dataList) {
        ExportState state = prepareContext(clazz, dataList);
        ExcelConfig config = clazz.getAnnotation(ExcelConfig.class);
        XSSFSheet sheet = createSheet(workbook, sheetName);
        writeHeader(sheet, dataList, config, state);
        writeData(sheet, dataList, state);
        applyMerges(sheet, dataList, state);
        applyColumnWidths(sheet);
    }

    /**
     * 解析默认 Sheet 名称。
     *
     * @param clazz 数据模型类型
     * @return Sheet 名称
     */
    private String resolveSheetName(Class<?> clazz) {
        ExcelConfig config = clazz.getAnnotation(ExcelConfig.class);
        return (config != null && !config.sheetName().isEmpty()) ? config.sheetName() : "Sheet1";
    }

    /**
     * 初始化一次导出所需的上下文和结构信息。
     *
     * @param clazz 数据模型类型
     * @param dataList 当前导出数据
     * @return 导出阶段状态
     */
    private ExportState prepareContext(Class<?> clazz, List<?> dataList) {
        ExcelContext context = ExcelContext.create(clazz, dataList);
        List<NestedProcessor.NestedFieldInfo> nestedFieldInfos = structureAnalyzer.collectNestedFieldInfos(context);
        List<ListFieldInfo> listFieldInfos = structureAnalyzer.collectListFieldInfos(context);
        int maxHeaderDepth = structureAnalyzer.calculateMaxHeaderDepth(context, nestedFieldInfos, listFieldInfos);
        context.setHeaderRows(maxHeaderDepth);
        context.setDataStartRow(maxHeaderDepth);
        return new ExportState(context, nestedFieldInfos, listFieldInfos, maxHeaderDepth);
    }

    /**
     * 创建工作表。
     *
     * @param workbook 工作簿
     * @param sheetName Sheet 名称
     * @return 工作表对象
     */
    private XSSFSheet createSheet(XSSFWorkbook workbook, String sheetName) {
        if (sheetName == null || sheetName.isEmpty()) {
            sheetName = "Sheet1";
        }
        return workbook.createSheet(sheetName);
    }

    /**
     * 写入表头区域。
     *
     * @param sheet 工作表
     * @param dataList 导出数据
     * @param config 导出配置
     * @param state 导出状态
     */
    private void writeHeader(XSSFSheet sheet, List<?> dataList, ExcelConfig config, ExportState state) {
        new HeaderWriter(exportSupport, sheet, dataList, config, state).write();
    }

    /**
     * 写入数据区域。
     *
     * @param sheet 工作表
     * @param dataList 导出数据
     * @param state 导出状态
     */
    private void writeData(XSSFSheet sheet, List<?> dataList, ExportState state) {
        new DataWriter(exportSupport, sheet, dataList, state).write();
    }

    /**
     * 应用数据区域合并策略。
     *
     * @param sheet 工作表
     * @param dataList 导出数据
     * @param state 导出状态
     */
    private void applyMerges(XSSFSheet sheet, List<?> dataList, ExportState state) {
        new MergeHandler(exportSupport, sheet, dataList, state).apply();
    }

    /**
     * 自动调整列宽。
     *
     * @param sheet 工作表
     */
    private void applyColumnWidths(XSSFSheet sheet) {
        int maxColumns = sheet.getRow(0) != null ? sheet.getRow(0).getLastCellNum() : 0;
        for (int i = 0; i < maxColumns; i++) {
            sheet.autoSizeColumn(i);
        }
    }
}
