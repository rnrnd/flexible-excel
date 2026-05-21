package com.flexibleexcel;

import com.flexibleexcel.annotation.Excel;
import com.flexibleexcel.annotation.ExcelConfig;
import com.flexibleexcel.core.ExcelExporter;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Excel导出测试
 */
public class ExcelExporterTest {

    private static final SimpleDateFormat TS = new SimpleDateFormat("HHmmss");

    private String ts() {
        return TS.format(new Date());
    }

    /**
     * 测试基本导出功能
     */
    @Test
    public void testBasicExport() {
        List<User> users = createTestData();
        String filePath = "target/test-basic-" + ts() + ".xlsx";

        ExcelExporter.create().export(User.class, users, filePath);

        System.out.println("基本导出完成: " + new File(filePath).getAbsolutePath());
    }

    /**
     * 测试List类型导出（纵向扩展）
     */
    @Test
    public void testListExport() {
        List<Order> orders = createOrderData();
        String filePath = "target/test-list-" + ts() + ".xlsx";

        ExcelExporter.create().export(Order.class, orders, filePath);

        System.out.println("List导出完成: " + new File(filePath).getAbsolutePath());
    }

    /**
     * 测试Map类型导出（横向扩展）
     */
    @Test
    public void testMapExport() {
        try {
            List<Product> products = createProductData();
            String filePath = "target/test-map-" + ts() + ".xlsx";

            ExcelExporter.create().export(Product.class, products, filePath);

            System.out.println("Map导出完成: " + new File(filePath).getAbsolutePath());
        } catch (Exception e) {
            System.err.println("Map导出失败: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    }

    /**
     * 测试复杂场景（同时包含List和Map）
     */
    @Test
    public void testComplexExport() {
        try {
            List<ComplexData> complexDataList = createComplexData();
            String filePath = "target/test-complex-" + ts() + ".xlsx";

            ExcelExporter.create().export(ComplexData.class, complexDataList, filePath);

            System.out.println("复杂导出完成: " + new File(filePath).getAbsolutePath());
        } catch (Exception e) {
            System.err.println("复杂导出失败: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    }

    /**
     * 测试嵌套对象导出（横向展开）
     */
    @Test
    public void testNestedHorizontalExport() {
        try {
            List<Student> students = createStudentData();
            String filePath = "target/test-nested-" + ts() + ".xlsx";

            ExcelExporter.create().export(Student.class, students, filePath);

            System.out.println("嵌套对象导出完成: " + new File(filePath).getAbsolutePath());
        } catch (Exception e) {
            System.err.println("嵌套对象导出失败: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    }

    /**
     * 测试复杂嵌套（同时包含List、Map和嵌套对象）
     */
    @Test
    public void testAllTypesExport() {
        try {
            List<SalesRecord> records = createSalesRecordData();
            String filePath = "target/test-all-types-" + ts() + ".xlsx";

            ExcelExporter.create().export(SalesRecord.class, records, filePath);

            System.out.println("全类型导出完成: " + new File(filePath).getAbsolutePath());
        } catch (Exception e) {
            System.err.println("全类型导出失败: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    }

    // ==================== 测试数据模型 ====================

    /**
     * 用户实体（基本字段）
     */
    @ExcelConfig(sheetName = "用户列表")
    @Data
    @AllArgsConstructor
    public static class User {
        @Excel(header = "姓名", columnWidth = 15, backgroundColor = "#E7E6E6", fontBold = true)
        private String name;

        @Excel(header = "年龄", columnWidth = 10)
        private Integer age;

        @Excel(header = "邮箱", columnWidth = 25)
        private String email;

        @Excel(header = "注册日期", columnWidth = 15, dateFormat = "yyyy-MM-dd")
        private Date registerDate;

        @Excel(header = "状态", columnWidth = 10, backgroundColor = "#D9EAD3")
        private String status;
    }

    /**
     * 订单实体（包含List）
     */
    @ExcelConfig(sheetName = "订单列表")
    @Data
    @AllArgsConstructor
    public static class Order {
        @Excel(header = "订单号", columnWidth = 20)
        private String orderNo;

        @Excel(header = "客户名称", columnWidth = 15)
        private String customerName;

        @Excel(header = "商品列表", columnWidth = 30, backgroundColor = "#FFF2CC")
        private List<String> items;

        @Excel(header = "金额", columnWidth = 12)
        private Double amount;
    }

    /**
     * 产品实体（包含Map）
     */
    @ExcelConfig(sheetName = "产品信息")
    @Data
    public static class Product {
        @Excel(header = "产品编号", columnWidth = 15)
        private String productNo;

        @Excel(header = "产品名称", columnWidth = 20)
        private String productName;

        @Excel(header = "规格参数", columnWidth = 40, backgroundColor = "#DAE8FC")
        private Map<String, String> specifications;

        @Excel(header = "价格", columnWidth = 12)
        private Double price;

        public Product(String productNo, String productName, Map<String, String> specifications, Double price) {
            this.productNo = productNo;
            this.productName = productName;
            this.specifications = specifications;
            this.price = price;
        }
    }

    /**
     * 复杂数据实体（同时包含List和Map）
     */
    @ExcelConfig(sheetName = "综合报表")
    @Data
    public static class ComplexData {
        @Excel(header = "编号", columnWidth = 12)
        private String id;

        @Excel(header = "标题", columnWidth = 20)
        private String title;

        @Excel(header = "标签", columnWidth = 25, backgroundColor = "#F4CCCC")
        private List<String> tags;

        @Excel(header = "属性", columnWidth = 50, backgroundColor = "#CFE2F3")
        private Map<String, Object> attributes;

        public ComplexData(String id, String title, List<String> tags, Map<String, Object> attributes) {
            this.id = id;
            this.title = title;
            this.tags = tags;
            this.attributes = attributes;
        }
    }

    // ==================== 测试数据生成 ====================

    private List<User> createTestData() {
        List<User> users = new ArrayList<>();
        users.add(new User("张三", 28, "zhangsan@example.com", new Date(), "活跃"));
        users.add(new User("李四", 35, "lisi@example.com", new Date(), "活跃"));
        users.add(new User("王五", 42, "wangwu@example.com", new Date(), "停用"));
        users.add(new User("赵六", 25, "zhaoliu@example.com", new Date(), "活跃"));
        return users;
    }

    private List<Order> createOrderData() {
        List<Order> orders = new ArrayList<>();

        List<String> items1 = Arrays.asList("iPhone 14", "AirPods Pro", "MagSafe充电器");
        orders.add(new Order("ORD20230001", "张三", items1, 12999.00));

        List<String> items2 = Arrays.asList("MacBook Pro 14\"", "Apple Pencil");
        orders.add(new Order("ORD20230002", "李四", items2, 19999.00));

        List<String> items3 = Arrays.asList("iPad Air", "Smart Keyboard", "鼠标", "USB-C转接线");
        orders.add(new Order("ORD20230003", "王五", items3, 8999.00));

        List<String> items4 = Collections.singletonList("Apple Watch Ultra");
        orders.add(new Order("ORD20230004", "赵六", items4, 5999.00));

        return orders;
    }

    private List<Product> createProductData() {
        List<Product> products = new ArrayList<>();

        Map<String, String> spec1 = new LinkedHashMap<>();
        spec1.put("颜色", "银色");
        spec1.put("存储", "256GB");
        spec1.put("屏幕", "6.1英寸");
        spec1.put("重量", "172g");
        products.add(new Product("P001", "iPhone 14", spec1, 5999.00));

        Map<String, String> spec2 = new LinkedHashMap<>();
        spec2.put("颜色", "深空灰");
        spec2.put("存储", "512GB");
        spec2.put("屏幕", "14.2英寸");
        spec2.put("重量", "1.6kg");
        products.add(new Product("P002", "MacBook Pro", spec2, 19999.00));

        Map<String, String> spec3 = new LinkedHashMap<>();
        spec3.put("颜色", "蓝色");
        spec3.put("尺寸", "40mm");
        spec3.put("材质", "钛金属");
        products.add(new Product("P003", "Apple Watch Ultra", spec3, 5999.00));

        return products;
    }

    private List<ComplexData> createComplexData() {
        List<ComplexData> list = new ArrayList<>();

        Map<String, Object> attrs1 = new LinkedHashMap<>();
        attrs1.put("材质", "铝合金");
        attrs1.put("尺寸", "310x240x25mm");
        attrs1.put("重量", "1.8kg");
        attrs1.put("保修", "2年");
        list.add(new ComplexData("C001", "笔记本电脑", Arrays.asList("电子产品", "办公", "便携"), attrs1));

        Map<String, Object> attrs2 = new LinkedHashMap<>();
        attrs2.put("材质", "硅胶");
        attrs2.put("颜色", "透明");
        attrs2.put("适用机型", "iPhone 14系列");
        list.add(new ComplexData("C002", "手机壳", Arrays.asList("配件", "保护"), attrs2));

        Map<String, Object> attrs3 = new LinkedHashMap<>();
        attrs3.put("容量", "10000mAh");
        attrs3.put("输入", "Type-C 18W");
        attrs3.put("输出", "USB-A 12W");
        attrs3.put("数量", "2个USB口");
        attrs3.put("特色", "双向快充");
        list.add(new ComplexData("C003", "充电宝", Arrays.asList("配件", "充电", "便携"), attrs3));

        return list;
    }

    // ==================== 嵌套对象测试模型 ====================

    /**
     * 地址信息（嵌套对象）
     */
    @AllArgsConstructor
    @Data
    public static class Address {
        @Excel(header = "省份")
        private String province;

        @Excel(header = "城市")
        private String city;

        @Excel(header = "区县")
        private String district;

        @Excel(header = "详细地址")
        private String detail;


    }

    /**
     * 联系信息（嵌套对象）
     */
    @Data
    @AllArgsConstructor
    public static class Contact {
        @Excel(header = "手机")
        private String phone;

        @Excel(header = "邮箱")
        private String email;

        @Excel(header = "QQ")
        private String qq;
    }

    /**
     * 学生实体（包含嵌套对象）
     */
    @ExcelConfig(sheetName = "学生信息")
    @Data
    @AllArgsConstructor
    public static class Student {
        @Excel(header = "学号", columnWidth = 12)
        private String studentNo;

        @Excel(header = "姓名", columnWidth = 10)
        private String name;

        @Excel(header = "地址", columnWidth = 50,
              nested = Excel.NestedMode.HORIZONTAL,
              nestedHeaderBgColor = "#4472C4",
              nestedSubHeaderBgColor = "#D6DCE5")
        private Address address;

        @Excel(header = "联系方式", columnWidth = 40,
              nested = Excel.NestedMode.HORIZONTAL,
              nestedHeaderBgColor = "#70AD47",
              nestedSubHeaderBgColor = "#E2EFDA")
        private Contact contact;
    }

    /**
     * 商品项（嵌套对象）
     */
    @Data
    @AllArgsConstructor
    public static class Item {
        @Excel(header = "商品名称")
        private String itemName;

        @Excel(header = "数量")
        private Integer quantity;

        @Excel(header = "单价")
        private Double unitPrice;

        public Double getTotalPrice() {
            return quantity * unitPrice;
        }
    }

    /**
     * 销售记录（同时包含List、Map和嵌套对象）
     */
    @ExcelConfig(sheetName = "销售记录")
    @Data
    @AllArgsConstructor
    public static class SalesRecord {
        @Excel(header = "记录ID", columnWidth = 12)
        private String recordId;

        @Excel(header = "销售人员", columnWidth = 12)
        private String salesperson;

        @Excel(header = "客户信息", columnWidth = 45,
              nested = Excel.NestedMode.HORIZONTAL,
              nestedHeaderBgColor = "#4472C4")
        private Customer customer;

        @Excel(header = "商品明细", columnWidth = 30, backgroundColor = "#FFF2CC")
        private List<Item> items;

        @Excel(header = "备注", columnWidth = 20)
        private String remarks;
    }

    /**
     * 客户信息（嵌套对象）
     */
    @Data
    @AllArgsConstructor
    public static class Customer {
        @Excel(header = "客户名称")
        private String customerName;

        @Excel(header = "联系电话")
        private String phone;

        @Excel(header = "收货地址")
        private String address;

        @Excel(header = "客户等级")
        private String level;
    }

    // ==================== 嵌套对象测试数据生成 ====================

    private List<Student> createStudentData() {
        List<Student> students = new ArrayList<>();

        students.add(new Student(
                "S001",
                "张三",
                new Address("北京市", "北京市", "朝阳区", "建国路88号SOHO现代城"),
                new Contact("13800138000", "zhangsan@example.com", "123456")
        ));

        students.add(new Student(
                "S002",
                "李四",
                new Address("上海市", "上海市", "浦东新区", "世纪大道100号环球金融中心"),
                new Contact("13900139000", "lisi@example.com", "654321")
        ));

        students.add(new Student(
                "S003",
                "王五",
                new Address("广东省", "深圳市", "南山区", "科技中一路腾讯大厦"),
                new Contact("13700137000", "wangwu@example.com", "111222")
        ));

        return students;
    }

    private List<SalesRecord> createSalesRecordData() {
        List<SalesRecord> records = new ArrayList<>();

        List<Item> items1 = new ArrayList<>();
        items1.add(new Item("iPhone 14", 2, 5999.00));
        items1.add(new Item("AirPods Pro", 2, 1999.00));
        records.add(new SalesRecord(
                "R001",
                "销售员A",
                new Customer("阿里巴巴", "400-800-0001", "浙江省杭州市余杭区", "VIP"),
                items1,
                "大客户订单"
        ));

        List<Item> items2 = new ArrayList<>();
        items2.add(new Item("MacBook Pro", 1, 19999.00));
        items2.add(new Item("Magic Mouse", 1, 699.00));
        items2.add(new Item("Magic Keyboard", 1, 999.00));
        records.add(new SalesRecord(
                "R002",
                "销售员B",
                new Customer("腾讯科技", "400-800-0002", "广东省深圳市南山区", "VIP"),
                items2,
                "企业采购"
        ));

        List<Item> items3 = new ArrayList<>();
        items3.add(new Item("iPad Air", 3, 4799.00));
        records.add(new SalesRecord(
                "R003",
                "销售员A",
                new Customer("字节跳动", "400-800-0003", "北京市海淀区", "普通"),
                items3,
                ""
        ));

        return records;
    }

    // ==================== 多Sheet导出测试 ====================

    /**
     * 测试多Sheet导出（使用链式API）
     */
    @Test
    public void testMultiSheetExport() {
        try {
            // 创建各Sheet数据
            List<User> users = createTestData();
            List<Order> orders = createOrderData();
            List<Student> students = createStudentData();

            String filePath = "target/test-multi-sheet-" + ts() + ".xlsx";

            // 使用链式API导出多个Sheet
            ExcelExporter.create()
                    .sheet("用户列表", User.class, users)
                    .sheet("订单列表", Order.class, orders)
                    .sheet("学生信息", Student.class, students)
                    .toFile(filePath);

            System.out.println("多Sheet导出完成: " + new File(filePath).getAbsolutePath());
        } catch (Exception e) {
            System.err.println("多Sheet导出失败: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    }

    /**
     * 测试多Sheet导出（使用类上的注解自动获取sheetName）
     */
    @Test
    public void testMultiSheetExportWithAnnotation() {
        try {
            List<User> users = createTestData();
            List<Order> orders = createOrderData();
            List<Product> products = createProductData();
            List<SalesRecord> salesRecords = createSalesRecordData();

            String filePath = "target/test-multi-sheet-annotation-" + ts() + ".xlsx";

            // 使用类上的@ExcelConfig注解的sheetName
            ExcelExporter.create()
                    .sheet(User.class, users)      // 使用 User 类的 @ExcelConfig
                    .sheet(Order.class, orders)     // 使用 Order 类的 @ExcelConfig
                    .sheet(Product.class, products) // 使用 Product 类的 @ExcelConfig
                    .sheet(SalesRecord.class, salesRecords) // 使用 SalesRecord 类的 @ExcelConfig
                    .toFile(filePath);

            System.out.println("多Sheet导出（注解）完成: " + new File(filePath).getAbsolutePath());
        } catch (Exception e) {
            System.err.println("多Sheet导出（注解）失败: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    }

    /**
     * 测试混合模式导出（单个导出和多Sheet导出）
     */
    @Test
    public void testMixedExport() {
        try {
            List<User> users = createTestData();
            List<SalesRecord> salesRecords = createSalesRecordData();

            String filePath = "target/test-mixed-" + ts() + ".xlsx";

            // 混合使用
            ExcelExporter.create()
                    .sheet("综合报表", User.class, users)
                    .sheet(SalesRecord.class, salesRecords)
                    .toFile(filePath);

            System.out.println("混合模式导出完成: " + new File(filePath).getAbsolutePath());
        } catch (Exception e) {
            System.err.println("混合模式导出失败: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    }
}
