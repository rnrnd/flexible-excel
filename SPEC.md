# Flexible Excel Export Maven Plugin

## 项目概述
开发一个Java Maven插件，用于灵活导出Excel文件。支持复杂数据结构（List、Map）的智能展开和单元格合并。

## 核心功能

### 1. @Excel 注解
在类的属性上添加@Excel注解来标记可导出字段，支持以下配置：
- `header()`: 主表头名称
- `columnWidth()`: 列宽（单位：字符数）
- `rowHeight()`: 行高（单位：点）
- `backgroundColor()`: 背景色（RGB格式）
- `fontColor()`: 字体颜色
- `fontBold()`: 是否加粗
- `horizontalAlign()`: 水平对齐方式
- `verticalAlign()`: 垂直对齐方式
- `dateFormat()`: 日期格式化

### 2. List 类型处理
- List属性沿着当前列纵向扩展，以多行形式展示数据
- **对于 `List<Item>` 类型（其中 Item 是带有 @Excel 注解的复杂对象）**：
  - 自动根据 Item 类的 @Excel 注解字段生成次级表头
  - 数据横向展开，Item 的每个字段占一列
  - 需要两级表头（主表头为 List 字段名，次级表头为 Item 的字段名）
  - 不进行纵向合并，每个 Item 的数据单独占一行
- 其他属性需要自适应纵向合并单元格（跨行合并）
- 支持嵌套List的递归处理

### 3. Map 类型处理
- Map内容横向扩展为多列展示
- Map的key作为次级表头
- Map的value作为单元格内容
- 主表头需要横向合并单元格（跨列合并）

### 4. 样式支持
- 表头样式自定义
- 数据单元格样式
- 交替行颜色
- 边框样式

## 技术栈
- Java 8+
- Apache POI 5.x
- Maven Plugin API
- JUnit 5 (测试)

## 项目结构
```
flexible-excel/
├── pom.xml
├── src/
│   ├── main/
│   │   └── java/
│   │       └── com/
│   │           └── flexibleexcel/
│   │               ├── annotation/
│   │               │   ├── Excel.java
│   │               │   ├── ExcelIgnore.java
│   │               │   └── ExcelConfig.java
│   │               ├── core/
│   │               │   ├── ExcelExporter.java
│   │               │   ├── ExcelContext.java
│   │               │   ├── CellData.java
│   │               │   └── CellStyle.java
│   │               ├── processor/
│   │               │   ├── FieldProcessor.java
│   │               │   ├── ListProcessor.java
│   │               │   └── MapProcessor.java
│   │               ├── util/
│   │               │   └── ReflectionUtil.java
│   │               └── exception/
│   │                   └── ExcelExportException.java
│   └── test/
│       └── java/
│           └── com/
│               └── flexibleexcel/
│                   └── ExcelExporterTest.java
└── README.md
```
