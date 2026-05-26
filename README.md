# Flexible Excel

灵活的Excel导出Maven库/插件，支持复杂数据结构的智能展开和单元格合并。

## 特性

- 📋 **注解驱动**：通过`@Excel`注解标记导出字段，零代码配置
- 📊 **List支持**：`List<String>`纵向展开为多行，其他列自动纵向合并；`List<复杂对象>`横向展开为多列子表头
- 🗂️ **Map支持**：`Map<K,V>`横向展开为多列，key为次级表头，主表头自动横向合并
- 🏗️ **嵌套对象**：支持嵌套对象的横向展开，且支持任意深度递归嵌套
- 📑 **多Sheet导出**：链式API一次导出多个Sheet
- 🎨 **样式丰富**：支持自定义表头、数据单元格样式（背景色、字体色、加粗、对齐等）
- 📏 **尺寸控制**：支持自定义列宽、行高
- 📅 **格式支持**：日期、数字格式化输出

## 快速开始

### 1. 添加依赖

```xml
<dependency>
    <groupId>io.github.rnrnd</groupId>
    <artifactId>flexible-excel</artifactId>
    <version>1.0.2</version>
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
| backgroundColor | String | "" | 背景色（如"#FF0000"） |
| fontColor | String | "" | 字体颜色 |
| fontBold | boolean | false | 是否加粗 |
| horizontalAlign | HorizontalAlign | CENTER | 水平对齐 |
| verticalAlign | VerticalAlign | CENTER | 垂直对齐 |
| dateFormat | String | "yyyy-MM-dd" | 日期格式 |
| numberFormat | String | "" | 数字格式（如"#,##0.00"） |
| order | int | 0 | 导出顺序 |
| nested | NestedMode | NONE | 嵌套模式：NONE / HORIZONTAL / RECURSIVE |
| nestedHeaderBgColor | String | "" | 嵌套主表头背景色 |
| nestedSubHeaderBgColor | String | "" | 嵌套次级表头背景色 |
| mergeNestedHeader | boolean | true | 是否合并嵌套主表头 |

### @ExcelConfig 注解

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| sheetName | String | "Sheet1" | 工作表名称 |
| headerRows | int | 1 | 表头行数（自动计算） |
| headerBackgroundColor | String | "#4472C4" | 表头背景色 |
| headerFontColor | String | "#FFFFFF" | 表头字体颜色 |
| headerFontBold | boolean | true | 表头是否加粗 |
| dataStartRow | int | 1 | 数据起始行 |
| dataStartColumn | int | 0 | 数据起始列 |

### @ExcelIgnore 注解

标记在字段上，该字段不会被导出。

## 高级用法

### List 类型（纵向扩展）

`List<String>` / `List<Integer>` 等简单类型List，每行数据纵向展开为多行，其他列自动纵向合并。

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

效果（3条数据，第一条有2件商品，第二条3件，第三条1件）：

| 订单号 | 客户 | 商品列表 | 金额 |
|--------|------|----------|------|
| ORD001 | 张三 | iPhone 14 | 5999 |
| | | AirPods Pro | |
| ORD002 | 李四 | MacBook | 12999 |
| | | Magic Mouse | |
| | | Magic Keyboard | |
| ORD003 | 王五 | iPad Air | 4799 |

### List<复杂对象>（横向展开次级表头）

当List的元素是复杂对象时，可以横向展开为多列子表头。

```java
public class Item {
    @Excel(header = "商品名称")
    private String itemName;

    @Excel(header = "数量")
    private Integer quantity;

    @Excel(header = "单价")
    private Double unitPrice;
}

public class SalesRecord {
    @Excel(header = "记录ID")
    private String recordId;

    @Excel(header = "商品明细")
    private List<Item> items;  // 复杂对象List，横向展开子列
}
```

效果（2行表头）：

| | | 商品明细 | | |
|---|---|---|---|---|
| 记录ID | ... | 商品名称 | 数量 | 单价 | ... |
| R001 | ... | iPhone 14 | 2 | 5999 | ... |
| | | AirPods Pro | 2 | 1999 | ... |

### Map 类型（横向扩展）

```java
public class Product {
    @Excel(header = "产品编号")
    private String productNo;

    @Excel(header = "产品名称")
    private String productName;

    @Excel(header = "规格参数")
    private Map<String, String> specs;  // 横向展开，key为子表头

    @Excel(header = "价格")
    private Double price;
}
```

效果：

| 产品编号 | 产品名称 | 颜色 | 存储 | 屏幕 | 重量 | 价格 |
|----------|----------|------|------|------|------|------|
| P001 | iPhone 14 | 银色 | 256GB | 6.1寸 | 172g | 5999 |
| P002 | MacBook | 深空灰 | 512GB | 14寸 | 1.6kg | 12999 |

### 嵌套对象（横向展开）

一个字段是复杂对象时，设置 `nested = HORIZONTAL` 将其属性横向展开为子列。

```java
public class Address {
    @Excel(header = "省份")
    private String province;

    @Excel(header = "城市")
    private String city;

    @Excel(header = "区县")
    private String district;

    @Excel(header = "详细地址")
    private String detail;
}

public class Student {
    @Excel(header = "学号")
    private String studentNo;

    @Excel(header = "姓名")
    private String name;

    @Excel(header = "地址",
          nested = Excel.NestedMode.HORIZONTAL,
          nestedHeaderBgColor = "#4472C4",
          nestedSubHeaderBgColor = "#D6DCE5")
    private Address address;
}
```

效果（2行表头）：

| | | 地址 | | | |
|---|---|---|---|---|---|
| 学号 | 姓名 | 省份 | 城市 | 区县 | 详细地址 |
| S001 | 张三 | 北京市 | 北京市 | 朝阳区 | 建国路88号 |
| S002 | 李四 | 上海市 | 上海市 | 浦东新区 | 世纪大道100号 |

### 多层嵌套对象（递归展开）

嵌套对象内部还可以继续包含嵌套对象，支持任意深度递归展开。

```java
// 联系方式详情
public class ContactDetail {
    @Excel(header = "手机")
    private String phone;

    @Excel(header = "邮箱")
    private String email;

    @Excel(header = "QQ")
    private String qq;
}

// 客户信息（包含嵌套联系方式）
public class DeepCustomer {
    @Excel(header = "客户名称")
    private String customerName;

    @Excel(header = "联系方式",
          nested = Excel.NestedMode.HORIZONTAL,
          nestedHeaderBgColor = "#70AD47")
    private ContactDetail contactDetail;  // 嵌套对象中的嵌套对象

    @Excel(header = "客户等级")
    private String level;
}

// 顶层模型
public class DeepNestedRecord {
    @Excel(header = "记录ID")
    private String recordId;

    @Excel(header = "客户信息",
          nested = Excel.NestedMode.HORIZONTAL,
          nestedHeaderBgColor = "#4472C4")
    private DeepCustomer customer;  // 顶层嵌套
}
```

效果（3行表头）：

| | | 客户信息 | | | | |
|---|---|---|---|---|---|---|
| | | | 联系方式 | | | |
| 记录ID | 销售人员 | 客户名称 | 手机 | 邮箱 | QQ | 客户等级 | 备注 |
| R001 | 销售员A | 阿里巴巴 | 138... | ali@... | 10001 | VIP | 大客户订单 |

> **说明**：只需在嵌套字段上继续标注 `@Excel(nested = Excel.NestedMode.HORIZONTAL)`，框架会自动递归展开。嵌套深度无限制。

### 多Sheet导出

使用链式API一次导出多个Sheet：

```java
ExcelExporter.create()
        .sheet("用户列表", User.class, users)
        .sheet("订单列表", Order.class, orders)
        .sheet("学生信息", Student.class, students)
        .toFile("report.xlsx");
```

也支持使用类上的 `@ExcelConfig(sheetName = "...")` 自动获取Sheet名称：

```java
ExcelExporter.create()
        .sheet(User.class, users)       // 使用 User 上 @ExcelConfig 的 sheetName
        .sheet(Order.class, orders)     // 使用 Order 上 @ExcelConfig 的 sheetName
        .sheet(Student.class, students)
        .toFile("report.xlsx");
```

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
│   ├── annotation/            # 注解定义
│   │   ├── Excel.java         # @Excel 主注解（含 NestedMode 枚举）
│   │   ├── ExcelConfig.java   # @ExcelConfig 配置注解
│   │   └── ExcelIgnore.java   # @ExcelIgnore 忽略注解
│   ├── core/                  # 核心类
│   │   ├── ExcelExporter.java # 导出器（链式API、递归表头/数据写入）
│   │   ├── ExcelContext.java  # 导出上下文（ExportField 解析）
│   │   ├── SheetData.java     # 多Sheet数据封装
│   │   ├── CellData.java      # 单元格数据模型
│   │   └── CellStyle.java     # 单元格样式
│   ├── processor/             # 字段处理器
│   │   ├── FieldProcessor.java  # 简单字段处理
│   │   ├── ListProcessor.java   # List字段处理
│   │   ├── MapProcessor.java    # Map字段处理
│   │   └── NestedProcessor.java # 嵌套对象处理（支持递归展开）
│   ├── util/                  # 工具类
│   │   └── ReflectionUtil.java  # 反射工具
│   └── exception/             # 异常
│       └── ExcelExportException.java
└── src/test/java/             # 测试用例
    └── ExcelExporterTest.java   # 包含基础、List、Map、嵌套、多层嵌套、多Sheet测试
```

## License

MIT License
