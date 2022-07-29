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

**读取单个sheet**

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