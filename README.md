> 阿里Easyexcel使用说明

# 什么Easyexcel？

Easyexcel是阿里工具开源的优秀的excel处理工具。

https://easyexcel.opensource.alibaba.com/

![image-20220729102715024](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207291027375.png)

# 使用教程

## 创建项目

使用IDEA创建一个Maven项目

![image-20220729103325518](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207291033607.png)

## 引入Easyexcel依赖

我使用的Easyexcel版本是3.1.1（目前最新版）

```xml
<?xml version="1.0" encoding="UTF-8"?>
<project xmlns="http://maven.apache.org/POM/4.0.0"
         xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
         xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 http://maven.apache.org/xsd/maven-4.0.0.xsd">
    <modelVersion>4.0.0</modelVersion>

    <groupId>org.itlab1024</groupId>
    <artifactId>easyexcel-tutorial</artifactId>
    <version>1.0-SNAPSHOT</version>

    <properties>
        <maven.compiler.source>17</maven.compiler.source>
        <maven.compiler.target>17</maven.compiler.target>
        <project.build.sourceEncoding>UTF-8</project.build.sourceEncoding>
    </properties>
    <dependencies>
        <dependency>
            <groupId>com.alibaba</groupId>
            <artifactId>easyexcel</artifactId>
            <version>3.1.1</version>
        </dependency>
        <!-- 工具类 -->
        <dependency>
            <groupId>com.alibaba</groupId>
            <artifactId>fastjson</artifactId>
            <version>2.0.10.graal</version>
        </dependency>
        <dependency>
            <groupId>org.projectlombok</groupId>
            <artifactId>lombok</artifactId>
            <version>1.18.24</version>
            <scope>compile</scope>
        </dependency>
        <dependency>
            <groupId>org.slf4j</groupId>
            <artifactId>slf4j-api</artifactId>
            <version>1.7.36</version>
        </dependency>
        <dependency>
            <groupId>ch.qos.logback</groupId>
            <artifactId>logback-core</artifactId>
            <version>1.2.11</version>
        </dependency>
        <dependency>
            <groupId>ch.qos.logback</groupId>
            <artifactId>logback-classic</artifactId>
            <version>1.2.11</version>
        </dependency>
        <dependency>
            <groupId>org.junit.jupiter</groupId>
            <artifactId>junit-jupiter</artifactId>
            <version>5.8.1</version>
            <scope>test</scope>
        </dependency>
    </dependencies>
</project>
```

## 读取Excel

### 基本读取

准备一个Excel文件，算上表头共有102条记录。

![image-20220729123525032](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207291235174.png)

定义跟excel表头一致的实体类

```java
package com.itlab1024.easyexcel.read;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

@Data
public class SampleData {
    @ExcelProperty(value = "姓名")
    private String name;
    @ExcelProperty(value = "年龄")
    private String age;
    @ExcelProperty(value = "出生年月")
    private Date birthday;
}
```

`@ExcelProperty(value = "姓名")`注解可以使用index来制定列的索引，或者使用value指定列表头来读取数据，可根据实际需要选择性设置，如果没有表头那就只能用索引了，如果有表头，不推荐使用索引，更不推荐在一个类中混合使用。

**还需要提醒的是**：如果不使用注解也是可以的，但是要保证类中字段的顺序要和excel列的顺序一致。

### 读取单个sheet

```java
@Test
public void testBasicRead() {
  EasyExcel.read("sample.xlsx", SampleData.class, new PageReadListener<SampleData>(dataList -> {
    log.info("读取到的数据是:{}", JSON.toJSONString(dataList));
  })).sheet().doRead();
}
```

说明：上面代码中使用了PageReadListener监听类，该类会每凑够100条数据，发送过来，比如我的excel种有101条数据（不包括表头），那么，上面log.info行代码会输出两次，第一次输出100条记录，第二条输出1条，结果如图所示

![image-20220729124111559](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207291241649.png)

我们可以可以自定义自己的监听器。只需要实现ReadListener接口，重写下面的方法即可。主要涉及两个方法`invoke`和`doAfterAllAnalysed`，顾名思义前者是监听器调度后处理数据，后者是解析完毕所有数据后的回调方法。

上面使用了`sheet()`方法，默认是读取第一个sheet，也可以传递名称或者index来指定sheet读取（多sheet读取其中的一个），index从0开始。



### 读取多个sheet

* 情况一：所有sheet数据格式统一（一类数据）

```java
@Test
public void testMultiSheetsRead() {
  EasyExcel.read("sample.xlsx", SampleData.class, new PageReadListener<SampleData>(dataList -> {
    log.info("读取到的数据是:{}", JSON.toJSONString(dataList));
  })).doReadAll();
}
```

使用`doReadAll`读取，并且只有一个监听器，也就是说所有sheet的数据都会向一个监听器中写，据我测试是按照sheet的顺序读取的数据。

* 情况二：读取多个sheet中的某几个。

```java
@Test
public void testMultiSheetsRead2() {
  ExcelReader excelReader = EasyExcel.read("sample.xlsx").build();
  // 比如我只读取前两个（两个是不同的格式）
  ReadSheet readSheet1 = EasyExcel.readSheet(0).head(SampleData.class).registerReadListener(new PageReadListener<>(dataList -> {
    log.info("readSheet1读取到的数据是:{}", JSON.toJSONString(dataList));
  })).build();
  ReadSheet readSheet2 = EasyExcel.readSheet(1).head(SampleData.class).registerReadListener(new PageReadListener<>(dataList -> {
    log.info("readSheet2读取到的数据是:{}", JSON.toJSONString(dataList));
  })).build();
  excelReader.read(readSheet1, readSheet2);
  excelReader.close();
}
```

运行结果是：

![image-20220729132046100](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207291320328.png)

注意：上面两个sheet我使用的是同一个类SampleData，但是实际情况可能是这两个sheet的格式不同，所有.head()需要传递不同的类，监听器也要使用不同的（如果不使用匿名监听器）。



### 格式转换

EasyExcel内置了日期和数字的格式转换，也支持自定义格式转换。

日期转化使用`@DateTimeFormat`注解：该注解仅可以在`java.util.Date`和`java.lang.String`两种类型上有效。

数字转化使用`@NumberFormat`注解：该注解仅可以在`java.lang.Number`和`java.lang.String`两种类型上有效，使用方法查看`java.text.DecimalFormat`类。

自定义转换类。

```java
package com.itlab1024.easyexcel.converter;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.converters.ReadConverterContext;
import com.alibaba.excel.converters.WriteConverterContext;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

public class NameConverter implements Converter<String> {
    @Override
    public Class<?> supportJavaTypeKey() {
        return String.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    /**
     * 读转换
     * @param context read converter context
     * @return
     * @throws Exception
     */
    @Override
    public String convertToJavaData(ReadConverterContext<?> context) throws Exception {
        String value = context.getReadCellData().getStringValue();
        if (null != value && value.contains("golang")) {
            return "已被转化";
        }
        return value;
    }

    /**
     * 写转换
     * @param context write context
     * @return
     * @throws Exception
     */
    @Override
    public WriteCellData<?> convertToExcelData(WriteConverterContext<String> context) throws Exception {
        return Converter.super.convertToExcelData(context);
    }
}
```

定义一个新的接收类

```java
package com.itlab1024.easyexcel.read;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.itlab1024.easyexcel.converter.NameConverter;
import lombok.Data;

import java.util.Date;

@Data
public class ConvertData {
    @ExcelProperty(value = "姓名", converter = NameConverter.class)
    private String name;
    @ExcelProperty(value = "年龄")
    @NumberFormat("#.##%")
    private String age;
    @ExcelProperty(value = "出生年月")
    @DateTimeFormat("yyyy年MM月dd日")
    private Date birthday;
}
```

测试方法

```java
@Test
public void testConvertRead() {
  EasyExcel.read("sample.xlsx", ConvertData.class, new PageReadListener<SampleData>(dataList -> {
    log.info("读取到的数据是:{}", JSON.toJSONString(dataList));
  })).sheet().doRead();
}
```

运行结果

![image-20220729134602590](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207291346703.png)



WTF：为什么生日没有被转化成功呢？将birthday的类型修改为String再次运行。

![image-20220729134659696](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207291346833.png)

不知道是我哪里用错了还是文档写错了？？？后续仔细查下。

### 同步读取

上面的都是异步读取，通过监听器处理。Easyexcel也提供了同步读取，同步读取有个弊端，大量数据会被放入到内存中。

```java
@Test
public void testSyncRead() {
  List<Object> objects = EasyExcel.read("sample.xlsx").head(SampleData.class).sheet().doReadSync();
  log.info("读取结果{}", JSON.toJSONString(objects));
}
```

运行结果

![image-20220729141208174](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207291412305.png)

### 读取表头

监听器中可以获取表头，监听器可以继承AnalysisEventListener类，也可以实现ReadListener接口。

继承AnalysisEventListener类：

```java
package com.itlab1024.easyexcel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.util.ConverterUtils;
import com.alibaba.fastjson2.JSON;
import com.itlab1024.easyexcel.read.SampleData;
import lombok.extern.slf4j.Slf4j;

import java.util.Map;

@Slf4j
public class GetHeadListener extends AnalysisEventListener<SampleData> {
    @Override
    public void invokeHead(Map<Integer, ReadCellData<?>> headMap, AnalysisContext context) {
        log.info("解析到一条头数据:{}", JSON.toJSONString(headMap));
        //转换结构
        Map<Integer, String> map = ConverterUtils.convertToStringMap(headMap, context);
        log.info("转换结构后的表头数据是{}", JSON.toJSONString(map));
    }

    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        super.invokeHeadMap(headMap, context);
        log.info("转换结构后的表头数据是{}", JSON.toJSONString(headMap));
    }

    @Override
    public void invoke(SampleData data, AnalysisContext context) {

    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }
}
```

这种方式可以实现两个方法，invokeHead和invokeHeadMap，两者都会返回表头的map，结构不同而已。后者返回的是Map<Integer, String>。更简洁，实际使用也更多。

---



实现ReadListener接口

该种方式只有invokeHead方法，如果想得到Map<Integer, String>类型的数据可以通过`ConverterUtils.convertToStringMap`转换。



获取表头测试代码如下：

```java
// 获取表头
@Test
public void testGetTableHeadRead() {
  EasyExcel.read("sample.xlsx", SampleData.class, new GetHeadListener()).sheet().doRead();
}
```

运行结果如下：

![image-20220729143001524](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207291430713.png)

### 读取批注，超链接，合并单元格

新增一个sheet，准备批注，超链接，合并单元格的数据。

![image-20220731142942452](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207311429803.png)

读取批注需要实现Listener中的如下方法

```java
package com.itlab1024.easyexcel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.read.listener.ReadListener;
import com.itlab1024.easyexcel.read.SampleData;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class ExtraListener implements ReadListener<SampleData> {
    @Override
    public void invoke(SampleData data, AnalysisContext context) {

    }

    @Override
    public void extra(CellExtra extra, AnalysisContext context) {
        CellExtraTypeEnum type = extra.getType();
        switch (type) {
            case COMMENT:
                log.info("额外信息是批注,在rowIndex:{},columnIndex;{},内容是:{}", extra.getRowIndex(), extra.getColumnIndex(),
                        extra.getText());
                break;
            case HYPERLINK:
                if ("Sheet1!A1".equals(extra.getText())) {
                    log.info("额外信息是超链接,在rowIndex:{},columnIndex;{},内容是:{}", extra.getRowIndex(),
                            extra.getColumnIndex(), extra.getText());
                } else if ("Sheet2!A1".equals(extra.getText())) {
                    log.info(
                            "额外信息是超链接,而且覆盖了一个区间,在firstRowIndex:{},firstColumnIndex;{},lastRowIndex:{},lastColumnIndex:{},"
                                    + "内容是:{}",
                            extra.getFirstRowIndex(), extra.getFirstColumnIndex(), extra.getLastRowIndex(),
                            extra.getLastColumnIndex(), extra.getText());
                } else {
                    log.info("超链接是:{}", extra.getText());
                }
                break;
            case MERGE:
                log.info(
                        "额外信息是单元格,在firstRowIndex:{},firstColumnIndex;{},lastRowIndex:{},lastColumnIndex:{}, 单元格内容是{}",
                        extra.getFirstRowIndex(), extra.getFirstColumnIndex(), extra.getLastRowIndex(),
                        extra.getLastColumnIndex(), extra.getText());
                break;
            default:
                break;
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }
}

```

执行结果如下：

```tex
15:06:56.675 [main] INFO com.itlab1024.easyexcel.listener.ExtraListener - 额外信息是单元格,在firstRowIndex:2,firstColumnIndex;0,lastRowIndex:2,lastColumnIndex:1, 单元格内容是null
15:06:56.683 [main] INFO com.itlab1024.easyexcel.listener.ExtraListener - 超链接是:https://itlab1024.com
15:06:56.719 [main] INFO com.itlab1024.easyexcel.listener.ExtraListener - 额外信息是批注,在rowIndex:1,columnIndex;0,内容是:itlab:这里有批注哦
```

### 读取公式和类型

准备数据

![image-20220731154518175](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207311545551.png)

定义接收类

```java
package com.itlab1024.easyexcel.read;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.data.CellData;
import lombok.Data;

@Data
public class CellDataType {
    @ExcelProperty("公式")
    private CellData<String> formula;
}
```

监听器

```java
package com.itlab1024.easyexcel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.itlab1024.easyexcel.read.CellDataType;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class CellTypeListener implements ReadListener<CellDataType> {
    @Override
    public void invoke(CellDataType data, AnalysisContext context) {
      log.info("类型是:{}", data);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }
}

```

测试类

```java
@Test
public void testCellDataTypeRead() {
  EasyExcel.read("sample.xlsx", CellDataType.class, new CellTypeListener()).sheet("公式").doRead();
}
```

运行结果

![image-20220731155011926](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207311550042.png)



### 数据转换异常处理

在监听器中有如下方法，用于异常处理。

```java
@Override
public void onException(Exception exception, AnalysisContext context) {}
```



### 不创建对象读

上面使用的都是创建接收类，也可以不创建对象读取Excel。数据会被放入到Map<Integer,String>中,看如下示例

```java
 /**
  * 不创建接收对象读取文件
  */
@Test
public void testNoModelDataRead() {
  EasyExcel.read("sample.xlsx",  new NoModelDataReadListener()).sheet().doRead();
}
```

 监听类

```java
package com.itlab1024.easyexcel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import lombok.extern.slf4j.Slf4j;

import java.util.Map;

@Slf4j
public class NoModelDataReadListener implements ReadListener<Map<Integer, String>> {

    @Override
    public void invoke(Map<Integer, String> data, AnalysisContext context) {
        log.info("读取到的数据信息是{}", data);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }
}
```

监听类中打印结果如下:

![image-20220731155935379](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202207311559597.png)

## 写入Excel

### 基本写入

创建数据类

```java
package com.itlab1024.easyexcel.write;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class WriteSampleData {
    @ExcelProperty("姓名")
    private String name;
    @ExcelProperty("年龄")
    private int age;
    @ExcelProperty("出生年月")
    private Date birthday;
}
```

准备数组数据。

```java
private static final List<WriteSampleData> sampleData = new ArrayList<>();

@BeforeAll
public static void initData() {
  for (int i = 0; i < 10; i++) {
    sampleData.add(new WriteSampleData("姓名" + i, i, new Date()));
  }
}
```

基本写入

```java
/**
 * 基本写入
 */
@Test
public void testBasicWrite() {
  EasyExcel.write("write.xlsx").sheet("基本写入").head(WriteSampleData.class).doWrite(sampleData);
}
```

执行结果：

![image-20220803094817019](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208030948310.png)



出生年月列宽比较窄导致无法正常显示，可以通过`@ColumnWidth(数值)`来设置。

修改后重新写入，执行结果如下：

![image-20220803095206297](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208030952416.png)

额~，我设置的100，有点大了😄。

还有其他写法，比如使用ExcelWriter等。我就不一一尝试了。

### 指定列、排除列写入

可以设置指定列或排除列的信息，来实现自由写入功能。

```java
/**
 * 指定列导出，排除列导出
 */
@Test
public void testIncludeExcludeWrite() {
  // 只导出姓名
  EasyExcel.write("write.xlsx").sheet("基本写入").head(WriteSampleData.class).includeColumnFieldNames(Collections.singleton("name")).doWrite(sampleData);
  // 不导出年龄
  EasyExcel.write("write2.xlsx").sheet("基本写入").head(WriteSampleData.class).excludeColumnFieldNames(Collections.singleton("age")).doWrite(sampleData);
}
```

执行结果如下：

![只导出姓名](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031000357.png)

![不导出年龄](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031001595.png)



### 列顺序设置

导出的Excel种列的顺序默认是根据类定义顺序一致，如果想调整顺序，除了调整类中的顺序外，可以通过index指定，index默认是0，如果中间有不指定的index，比如设置了0设置了2，未设置1，则第二列是空列。

```java
package com.itlab1024.easyexcel.write;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class WriteSampleDataIndexed {
    @ExcelProperty(value = "姓名", index = 1)
    private String name;
    @ExcelProperty(value = "年龄", index = 2)
    private int age;
    @ExcelProperty(value = "出生年月", index = 0)
    @ColumnWidth(100)
    private Date birthday;
}

```

运行结果：

![image-20220803100518726](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031005816.png)



### 复杂表头

Easyexcel通过使用`@ExcelProperty({"主标题", "字符串标题"})`来设置负载表头。

比如如下配置

```java
package com.itlab1024.easyexcel.write;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class WriteSampleDataComplexHeader {
    @ExcelProperty({"基本信息", "姓名"})
    private String name;
    @ExcelProperty(value = {"基本信息", "年龄"})
    private int age;
    @ExcelProperty(value = {"出生年月"})
    @ColumnWidth(30)
    private Date birthday;
}
```

写入代码

```java
/**
 * 复杂表头
 */
@Test
public void testComplexHeaderWrite() {
  EasyExcel.write("write.xlsx").sheet("基本写入").head(WriteSampleDataComplexHeader.class).doWrite(sampleData);
}
```



运行结果：

![image-20220803101333264](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031013370.png)



### 重复多次写入

这在数据量很大的时候非常有用。比如一个excel数据可能有上百万行数据，如果数据一次性加载到内存可能会非常大，造成内存溢出。

重复多次写入主要通过ExcelWriter类。

```java
/**
 * 重复多次写入，比如我有三十条数据分三次写入到一个sheet中。
 */
@Test
public void testRepeatWrite() {
  ExcelWriter excelWriter = EasyExcel.write("write.xlsx", WriteSampleData.class).build();
  WriteSheet writeSheet = EasyExcel.writerSheet("重复多次写入").build();
  //模拟写入30条数据，每次写入10条数据
  for (int i = 0; i < 3; i++) {
    excelWriter.write(sampleData, writeSheet);
  }
  excelWriter.close();
}
```

![image-20220803102332154](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031023283.png)



也可以写入多个sheet中，创建writeSheet对象。多个我就不尝试了，粘贴下官网的例子

```java
try (ExcelWriter excelWriter = EasyExcel.write(fileName, DemoData.class).build()) {
// 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来。这里最终会写到5个sheet里面
for (int i = 0; i < 5; i++) {
// 每次都要创建writeSheet 这里注意必须指定sheetNo 而且sheetName必须不一样
WriteSheet writeSheet = EasyExcel.writerSheet(i, "模板" + i).build();
// 分页去数据库查询数据 这里可以去数据库查询每一页的数据
List<DemoData> data = data();
excelWriter.write(data, writeSheet);
}
}
```

也可以写入不同的sheet，并且数据不同，也就是header不同。

```java
// 方法3 如果写到不同的sheet 不同的对象
fileName = TestFileUtil.getPath() + "repeatedWrite" + System.currentTimeMillis() + ".xlsx";
// 这里 指定文件
try (ExcelWriter excelWriter = EasyExcel.write(fileName).build()) {
  // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来。这里最终会写到5个sheet里面
  for (int i = 0; i < 5; i++) {
    // 每次都要创建writeSheet 这里注意必须指定sheetNo 而且sheetName必须不一样。这里注意DemoData.class 可以每次都变，我这里为了方便 所以用的同一个class
    // 实际上可以一直变
    WriteSheet writeSheet = EasyExcel.writerSheet(i, "模板" + i).head(DemoData.class).build();
    // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
    List<DemoData> data = data();
    excelWriter.write(data, writeSheet);
  }
}
```

### 自定义写入Excel列的格式

跟之前介绍的coverter类似。

```java
@ExcelProperty(value = "字符串标题", converter = CustomStringStringConverter.class)
```

这里要重写的是如下方法

```java
@Override
public WriteCellData<?> convertToExcelData(WriteConverterContext<String> context) throws Exception {
  return Converter.super.convertToExcelData(context);
}
```

### 图片写入

图片写入支持多种类型。我就根据官网的例子尝试下。

定义支持的多种图片写入方式

```java
package com.itlab1024.easyexcel.write;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.converters.string.StringImageConverter;
import com.alibaba.excel.metadata.data.WriteCellData;
import lombok.Data;

import java.io.File;
import java.io.InputStream;
import java.net.URL;

@Data
public class WriteImageSampleData {
    private File file;
    private InputStream inputStream;
    /**
     * 如果string类型 必须指定转换器，string默认转换成string
     */
    @ExcelProperty(converter = StringImageConverter.class)
    private String string;
    private byte[] byteArray;
    /**
     * 根据url导出
     *
     * @since 2.1.1
     */
    private URL url;

    /**
     * 根据文件导出 并设置导出的位置。
     *
     * @since 3.0.0-beta1
     */
    private WriteCellData<Void> writeCellDataFile;

    public WriteImageSampleData() {
    }
}
```

写入图片代码

```java
/**
 * 写入图片
 * @throws Exception
 */
@Test
public void testImageWrite() throws Exception {
  WriteImageSampleData imageSampleData = new WriteImageSampleData();
  File file = new File("image.jpeg");
  imageSampleData.setFile(file);
  imageSampleData.setUrl(new URL("https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031034134.jpeg"));
  imageSampleData.setByteArray(IOUtils.toByteArray(new FileInputStream(file)));
  imageSampleData.setInputStream(new FileInputStream(file));
  imageSampleData.setString("image.jpeg");
  WriteCellData<Void> cellData = new WriteCellData<>();
  List<ImageData> imageDataList = new ArrayList<>();
  ImageData imageData = new ImageData();
  imageDataList.add(imageData);
  cellData.setImageDataList(imageDataList);
  // 放入2进制图片
  imageData.setImage(FileUtils.readFileToByteArray(new File("image.jpeg")));
  // 图片类型
  imageData.setImageType(ImageData.ImageType.PICTURE_TYPE_PNG);
  // 上 右 下 左 需要留空
  // 这个类似于 css 的 margin
  // 这里实测 不能设置太大 超过单元格原始大小后 打开会提示修复。暂时未找到很好的解法。
  imageData.setTop(5);
  imageData.setRight(40);
  imageData.setBottom(5);
  imageData.setLeft(5);
  cellData.setImageDataList(imageDataList);
  imageSampleData.setWriteCellDataFile(cellData);
  EasyExcel.write("write.xlsx", WriteImageSampleData.class).sheet().doWrite(Collections.singleton(imageSampleData));
}
```



运行结果



![image-20220803105623915](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031056075.png)

### 超链接、备注、公式、样式等设置方式

(⊙o⊙)…偷懒了，这块忽略了，附上官方代码（官方代码有些地方也是错的。。。。。。）

```java
 /**
  * 超链接、备注、公式、指定单个单元格的样式、单个单元格多种样式
  * <p>
  * 1. 创建excel对应的实体对象 参照{@link WriteCellDemoData}
  * <p>
  * 2. 直接写即可
  *
  * @since 3.0.0-beta1
  */
@Test
public void writeCellDataWrite() {
  String fileName = TestFileUtil.getPath() + "writeCellDataWrite" + System.currentTimeMillis() + ".xlsx";
  WriteCellDemoData writeCellDemoData = new WriteCellDemoData();

  // 设置超链接
  WriteCellData<String> hyperlink = new WriteCellData<>("官方网站");
  writeCellDemoData.setHyperlink(hyperlink);
  HyperlinkData hyperlinkData = new HyperlinkData();
  hyperlink.setHyperlinkData(hyperlinkData);
  hyperlinkData.setAddress("https://github.com/alibaba/easyexcel");
  hyperlinkData.setHyperlinkType(HyperlinkType.URL);

  // 设置备注
  WriteCellData<String> comment = new WriteCellData<>("备注的单元格信息");
  writeCellDemoData.setCommentData(comment);
  CommentData commentData = new CommentData();
  comment.setCommentData(commentData);
  commentData.setAuthor("Jiaju Zhuang");
  commentData.setRichTextStringData(new RichTextStringData("这是一个备注"));
  // 备注的默认大小是按照单元格的大小 这里想调整到4个单元格那么大 所以向后 向下 各额外占用了一个单元格
  commentData.setRelativeLastColumnIndex(1);
  commentData.setRelativeLastRowIndex(1);

  // 设置公式
  WriteCellData<String> formula = new WriteCellData<>();
  writeCellDemoData.setFormulaData(formula);
  FormulaData formulaData = new FormulaData();
  formula.setFormulaData(formulaData);
  // 将 123456789 中的第一个数字替换成 2
  // 这里只是例子 如果真的涉及到公式 能内存算好尽量内存算好 公式能不用尽量不用
  formulaData.setFormulaValue("REPLACE(123456789,1,1,2)");

  // 设置单个单元格的样式 当然样式 很多的话 也可以用注解等方式。
  WriteCellData<String> writeCellStyle = new WriteCellData<>("单元格样式");
  writeCellStyle.setType(CellDataTypeEnum.STRING);
  writeCellDemoData.setWriteCellStyle(writeCellStyle);
  WriteCellStyle writeCellStyleData = new WriteCellStyle();
  writeCellStyle.setWriteCellStyle(writeCellStyleData);
  // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.
  writeCellStyleData.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
  // 背景绿色
  writeCellStyleData.setFillForegroundColor(IndexedColors.GREEN.getIndex());

  // 设置单个单元格多种样式
  WriteCellData<String> richTest = new WriteCellData<>();
  richTest.setType(CellDataTypeEnum.RICH_TEXT_STRING);
  writeCellDemoData.setRichText(richTest);
  RichTextStringData richTextStringData = new RichTextStringData();
  richTest.setRichTextStringDataValue(richTextStringData);
  richTextStringData.setTextString("红色绿色默认");
  // 前2个字红色
  WriteFont writeFont = new WriteFont();
  writeFont.setColor(IndexedColors.RED.getIndex());
  richTextStringData.applyFont(0, 2, writeFont);
  // 接下来2个字绿色
  writeFont = new WriteFont();
  writeFont.setColor(IndexedColors.GREEN.getIndex());
  richTextStringData.applyFont(2, 4, writeFont);

  List<WriteCellDemoData> data = new ArrayList<>();
  data.add(writeCellDemoData);
  EasyExcel.write(fileName, WriteCellDemoData.class).inMemory(true).sheet("模板").doWrite(data);
}
```

### 注解的使用

**列宽、行高**

行高主要有表头的行高和内容的行高，分别是如下两个注解，是能使用在类上。

```java
@ContentRowHeight(10)
@HeadRowHeight(20)
```

列宽使用如下注解

```java
@ColumnWidth(50)
```

**字体风格**

```java
@ContentFontStyle
```

**表头样式（对齐等等）**

```
@HeadStyle
```



**合并单元格**



```java
// 这一列 每隔2行 合并单元格
@ContentLoopMerge(eachRow = 2)


// 将第6-7行的2-3列合并成一个单元格
@OnceAbsoluteMerge(firstRowIndex = 5, lastRowIndex = 6, firstColumnIndex = 1, lastColumnIndex = 2)
```

示例代码：

```java
package com.itlab1024.easyexcel.write;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.*;
import com.alibaba.excel.enums.BooleanEnum;
import com.alibaba.excel.enums.poi.HorizontalAlignmentEnum;
import com.alibaba.excel.enums.poi.VerticalAlignmentEnum;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Font;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
@HeadRowHeight(30) // 表头行高
@ContentRowHeight(50) // 内容行高
@ColumnWidth(30) // 列宽
@ContentFontStyle(fontName = "monaco", bold = BooleanEnum.TRUE, color = Font.COLOR_RED, underline = Font.U_DOUBLE) // 内容文字风格
@HeadFontStyle(fontName = "Arial", bold = BooleanEnum.TRUE, color = Font.COLOR_RED, underline = Font.U_SINGLE_ACCOUNTING) // 表头文字风格
@HeadStyle(horizontalAlignment = HorizontalAlignmentEnum.LEFT, verticalAlignment = VerticalAlignmentEnum.CENTER) //表头风格
@OnceAbsoluteMerge(firstRowIndex = 5, lastRowIndex = 6, firstColumnIndex = 1, lastColumnIndex = 2)
public class WriteSampleDataAnnotation {
    @ExcelProperty("姓名")
    @ContentLoopMerge(eachRow = 2)
    private String name;
    @ExcelProperty("年龄")
    private int age;
    @ExcelProperty("出生年月")
    @ColumnWidth(50) // 单独设置 birthday列宽
    private Date birthday;
}
```



运行结果：

![image-20220803122051934](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031220078.png)



### 使用Table写入

主要使用writerTable()方法

```java
@Test
public void testTableWrite() {
  WriteTable writeTable = EasyExcel.writerTable()
    .needHead(Boolean.TRUE) // 是否需要表头
    .tableNo(0) // 表索引
    .build();
  ExcelWriter excelWriter = EasyExcel.write("write.xlsx").build();
  WriteSheet writeSheet = EasyExcel.writerSheet("Table写入").build();
  excelWriter.write(sampleData, writeSheet, writeTable);
  excelWriter.close();
}
```



![image-20220803122735404](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031227588.png)



### 动态表头

动态表头主要是通过.head()方法设置，传递一个List<List<String>>类型的参数。

```java
@Test
public void testDynamicHeadWrite() {
  EasyExcel.write("write.xlsx")
    .head(makeHead()).sheet("动态表头")
    .doWrite(sampleData);
}

private List<List<String>> makeHead() {
  List<List<String>> lists = new ArrayList<>();
  List<String> list = Lists.newArrayList();
  list.add("合并表头");
  list.add("姓名");
  List<String> list2 = Lists.newArrayList();
  list2.add("合并表头");
  list2.add("年龄");
  List<String> list3 = Lists.newArrayList();
  list3.add("出生年月");
  lists.add(list);
  lists.add(list2);
  lists.add(list3);
  return lists;
}
```

运行结果：

![image-20220803123333485](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031233667.png)

### 自动列宽

自动列宽是源码内置的一个handler处理器，使用方法如下。

```java
@Test
public void  testAutoCellWidthWrite() {
  EasyExcel.write("write.xlsx", WriteSampleData.class)
    .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy()).sheet("模板").doWrite(sampleData);
}
```

运行结果：

不设置自动列宽

![image-20220803123950568](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031239659.png)

设置自动列宽后

![image-20220803123847610](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031238748.png)



### 自定义拦截器

以上方式都不能满足需求的情况下，可以自定义拦截器实现。

官网demo代码如下:

```java
/**
 * 自定义拦截器。对第一行第一列的头超链接到:https://github.com/alibaba/easyexcel
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class CustomCellWriteHandler implements CellWriteHandler {

    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        Cell cell = context.getCell();
        // 这里可以对cell进行任何操作
        log.info("第{}行，第{}列写入完成。", cell.getRowIndex(), cell.getColumnIndex());
        if (BooleanUtils.isTrue(context.getHead()) && cell.getColumnIndex() == 0) {
            CreationHelper createHelper = context.getWriteSheetHolder().getSheet().getWorkbook().getCreationHelper();
            Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
            hyperlink.setAddress("https://github.com/alibaba/easyexcel");
            cell.setHyperlink(hyperlink);
        }
    }

}


/**
 * 自定义拦截器.对第一列第一行和第二行的数据新增下拉框，显示 测试1 测试2
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class CustomSheetWriteHandler implements SheetWriteHandler {

    @Override
    public void afterSheetCreate(SheetWriteHandlerContext context) {
        log.info("第{}个Sheet写入成功。", context.getWriteSheetHolder().getSheetNo());

        // 区间设置 第一列第一行和第二行的数据。由于第一行是头，所以第一、二行的数据实际上是第二三行
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(1, 2, 0, 0);
        DataValidationHelper helper = context.getWriteSheetHolder().getSheet().getDataValidationHelper();
        DataValidationConstraint constraint = helper.createExplicitListConstraint(new String[] {"测试1", "测试2"});
        DataValidation dataValidation = helper.createValidation(constraint, cellRangeAddressList);
        context.getWriteSheetHolder().getSheet().addValidationData(dataValidation);
    }
}


 /**
     * 下拉，超链接等自定义拦截器（上面几点都不符合但是要对单元格进行操作的参照这个）
     * <p>
     * demo这里实现2点。1. 对第一行第一列的头超链接到:https://github.com/alibaba/easyexcel 2. 对第一列第一行和第二行的数据新增下拉框，显示 测试1 测试2
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link DemoData}
     * <p>
     * 2. 注册拦截器 {@link CustomCellWriteHandler} {@link CustomSheetWriteHandler}
     * <p>
     * 2. 直接写即可
     */
    @Test
    public void customHandlerWrite() {
        String fileName = TestFileUtil.getPath() + "customHandlerWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName, DemoData.class).registerWriteHandler(new CustomSheetWriteHandler())
            .registerWriteHandler(new CustomCellWriteHandler()).sheet("模板").doWrite(data());
    }

```

## 填充Excel

填充Excel主要是通过在excel中使用{}定义需要替换掉问题，比如{name}就对应类中的name属性。

预置一个模板文件。

![image-20220803125456665](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031254786.png)



### 基本写入

模板写入，

> 特别注意：数据类无需使用@ExcelProperty等注解（因为文件已经有了，就是替换嘛）,加上反而有问题，我遇到的问题是程序没有出错，但是数据未成功写入。

代码如下

```
package com.itlab1024.easyexcel.write;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class WriteTemplateSampleData {
    private String name;
    private int age;
    private Date birthday;
}

```

写入代码

```java
@Test
public void testTemplateBasicWrite() {
  WriteTemplateSampleData writeTemplateSampleData = new WriteTemplateSampleData();
  writeTemplateSampleData.setName("张三");
  writeTemplateSampleData.setAge(5);
  writeTemplateSampleData.setBirthday(new Date());
  EasyExcel.write("write.xlsx").withTemplate("template.xlsx").sheet().doFill(writeTemplateSampleData);
}
```

运行结果：

![image-20220803125629544](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031256690.png)



### 填充列表

上面的李四是一条数据，如果是多条数据如何填充呢？这就得修改模板使用{.属性名}来实现。

![image-20220803125943146](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031259288.png)

代码如下

```java
@Test
public void testTemplateListWrite() {
  WriteTemplateSampleData writeTemplateSampleData = new WriteTemplateSampleData();
  writeTemplateSampleData.setName("张三");
  writeTemplateSampleData.setAge(5);
  writeTemplateSampleData.setBirthday(new Date());
  WriteTemplateSampleData writeTemplateSampleData2 = new WriteTemplateSampleData();
  writeTemplateSampleData2.setName("张三2");
  writeTemplateSampleData2.setAge(5);
  writeTemplateSampleData2.setBirthday(new Date());
  List<WriteTemplateSampleData> datas = new ArrayList<>();
  datas.add(writeTemplateSampleData2);
  datas.add(writeTemplateSampleData);
  EasyExcel.write("write.xlsx").withTemplate("templateList.xlsx").sheet().doFill(datas);
  //
  // 方案2 分多次 填充 会使用文件缓存（省内存） jdk8
  // since: 3.0.0-beta1
  //        EasyExcel.write("write.xlsx").withTemplate("templateList.xlsx").sheet()
  //                .doFill(() -> {
  //                    // 分页查询数据
  //                    return null;
  //                });

  // 方案3 分多次 填充 会使用文件缓存（省内存）
  //        try (ExcelWriter excelWriter = EasyExcel.write("write.xlsx").withTemplate("templateList.xlsx").build()) {
  //            WriteSheet writeSheet = EasyExcel.writerSheet().build();
  //            excelWriter.fill(分片数据, writeSheet);
  //            excelWriter.fill(分片数据, writeSheet);
  //        }
}
```



执行结果：

![image-20220803130515309](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031305476.png)

# 复杂填充

当excel表比较复杂，比如数据分很多类的时候，可以将数据使用Map这种数据类型包装起来。代码与上面类似。

### 横向填充

有时候可能是一个合并的单元格，第一列是统计，后面有好几列，分别是子统计，比如如下图

![下载](https://itlab1024-1256529903.cos.ap-beijing.myqcloud.com/202208031317499.png)

就可以使用FillConfig的direction(WriteDirectionEnum.HORIZONTAL)方法配置

示例如下代码

```java
 // 方案1
try (ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build()) {
  WriteSheet writeSheet = EasyExcel.writerSheet().build();
  FillConfig fillConfig = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
  excelWriter.fill(data(), fillConfig, writeSheet);
  excelWriter.fill(data(), fillConfig, writeSheet);

  Map<String, Object> map = new HashMap<>();
  map.put("date", "2019年10月9日13:28:28");
  excelWriter.fill(map, writeSheet);
}
```

> 总结: Easyexcel还是一个非常不错的工具，内存确实得到了优化，对于大数据Excel的导入导出很棒！
>
> 后面有点偷懒了，并没有一一尝试。
