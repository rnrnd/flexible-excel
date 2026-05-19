# Flexible Excel

灵活的Excel导出Maven库/插件，支持复杂数据结构的智能展开和单元格合并。

## 特性

- 📋 **注解驱动**：通过`@Excel`注解标记导出字段
- 📊 **List支持**：List类型属性纵向展开为多行，其他列自动纵向合并
- 🗂️ **Map支持**：Map类型横向展开为多列，key为次级表头，主表头自动横向合并
- 🎨 **样式丰富**：支持自定义表头、数据单元格样式（背景色、字体色、加粗、对齐等）
- 📏 **尺寸控制**：支持自定义列宽、行高
- 📅 **格式支持**：日期、数字格式化输出

## 快速开始

### 1. 添加依赖

```xml
<dependency>
    <groupId>com.flexibleexcel</groupId>
    <artifactId>flexible-excel</artifactId>
    <version>1.0.0</version>
</dependency>
```

### 2. 定义数据模型

```java
@ExcelConfig(sheetName = "用户列表")
public class User {
    @Excel(header = "姓名", columnWidth = 15)
    private String name;

    @Excel(header = "年龄", columnWidth = 10)
    private Integer age;

    @Excel(header = "邮箱", columnWidth = 25)
    private String email;

    @Excel(header = "注册日期", dateFormat = "yyyy-MM-dd")
    private Date registerDate;
}
```

### 3. 执行导出

```java
List<User> users = Arrays.asList(
    new User("张三", 28, "zhangsan@example.com", new Date()),
    new User("李四", 35, "lisi@example.com", new Date())
);

ExcelExporter.create().export(User.class, users, "users.xlsx");
```

## 注解说明

### @Excel 注解

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| header | String | "" | 主表头名称 |
| columnWidth | int | -1 | 列宽（字符数，-1表示自动） |
| rowHeight | int | -1 | 行高（点） |
| backgroundColor | String | "" | 背景色（RGB格式，如"#FF0000"） |
| fontColor | String | "" | 字体颜色 |
| fontBold | boolean | false | 是否加粗 |
| horizontalAlign | HorizontalAlign | CENTER | 水平对齐 |
| verticalAlign | VerticalAlign | CENTER | 垂直对齐 |
| dateFormat | String | "yyyy-MM-dd" | 日期格式 |
| numberFormat | String | "" | 数字格式 |
| order | int | 0 | 导出顺序 |

### @ExcelConfig 注解

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| sheetName | String | "Sheet1" | 工作表名称 |
| headerRows | int | 1 | 表头行数 |
| headerBackgroundColor | String | "#4472C4" | 表头背景色 |
| headerFontColor | String | "#FFFFFF" | 表头字体颜色 |
| headerFontBold | boolean | true | 表头是否加粗 |
| dataStartRow | int | 1 | 数据起始行 |
| dataStartColumn | int | 0 | 数据起始列 |

## 高级用法

### List 类型（纵向扩展）

```java
public class Order {
    @Excel(header = "订单号")
    private String orderNo;

    @Excel(header = "客户")
    private String customerName;

    @Excel(header = "商品列表")
    private List<String> items;  // 纵向展开

    @Excel(header = "金额")
    private Double amount;
}
```

效果：
| 订单号 | 客户 | 商品列表 | 金额 |
|--------|------|----------|------|
| ORD001 | 张三 | iPhone 14 | 5999 |
| | | AirPods | 1999 |
| ORD002 | 李四 | MacBook | 12999 |

### Map 类型（横向扩展）

```java
public class Product {
    @Excel(header = "产品编号")
    private String productNo;

    @Excel(header = "产品名称")
    private String productName;

    @Excel(header = "规格参数")
    private Map<String, String> specs;  // 横向展开

    @Excel(header = "价格")
    private Double price;
}
```

效果：
| 产品编号 | 产品名称 | 颜色 | 存储 | 屏幕 | 价格 |
|----------|----------|------|------|------|------|
| P001 | iPhone 14 | 银色 | 256GB | 6.1寸 | 5999 |
| P002 | MacBook | 深空灰 | 512GB | 14寸 | 12999 |

## 安装到本地仓库

```bash
mvn clean install
```

## 运行测试

```bash
mvn test
```

## 项目结构

```
flexible-excel/
├── src/main/java/com/flexibleexcel/
│   ├── annotation/          # 注解定义
│   │   ├── Excel.java       # 主注解
│   │   ├── ExcelConfig.java # 配置注解
│   │   └── ExcelIgnore.java # 忽略注解
│   ├── core/                # 核心类
│   │   ├── ExcelExporter.java  # 导出器
│   │   ├── ExcelContext.java   # 上下文
│   │   ├── CellData.java       # 单元格数据
│   │   └── CellStyle.java      # 单元格样式
│   ├── processor/          # 处理器
│   │   ├── FieldProcessor.java # 字段处理
│   │   ├── ListProcessor.java # List处理
│   │   └── MapProcessor.java  # Map处理
│   ├── util/               # 工具类
│   │   └── ReflectionUtil.java
│   └── exception/          # 异常
│       └── ExcelExportException.java
└── src/test/java/          # 测试
```

## License

MIT License
