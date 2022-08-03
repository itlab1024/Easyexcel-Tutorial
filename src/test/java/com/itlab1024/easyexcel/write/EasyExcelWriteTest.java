package com.itlab1024.easyexcel.write;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.data.ImageData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.util.FileUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.WriteTable;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.itlab1024.easyexcel.read.SampleData;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.util.IOUtils;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

public class EasyExcelWriteTest {
    private static final List<WriteSampleData> sampleData = new ArrayList<>();

    @BeforeAll
    public static void initData() {
        for (int i = 0; i < 10; i++) {
            sampleData.add(new WriteSampleData("姓名" + i, i, new Date()));
        }
    }

    /**
     * 基本写入
     */
    @Test
    public void testBasicWrite() {
        EasyExcel.write("write.xlsx").sheet("基本写入").head(WriteSampleData.class).doWrite(sampleData);
    }


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

    /**
     * 设置Excel的列顺序
     */
    @Test
    public void testIndexedWrite() {
        EasyExcel.write("write.xlsx").sheet("基本写入").head(WriteSampleDataIndexed.class).doWrite(sampleData);
    }


    /**
     * 复杂表头
     */
    @Test
    public void testComplexHeaderWrite() {
        EasyExcel.write("write.xlsx").sheet("基本写入").head(WriteSampleDataComplexHeader.class).doWrite(sampleData);
    }

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
        EasyExcel.write("write.xlsx", WriteSampleData.class).sheet().doWrite(Collections.singleton(imageSampleData));
    }

    /**
     * 注解
     */
    @Test
    public void testAnnotationWrite() {
        EasyExcel.write("write.xlsx").sheet("基本写入").head(WriteSampleDataAnnotation.class).doWrite(sampleData);
    }

    /**
     * 表格方式写入
     */
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

    /**
     * 动态表头
     */
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

    @Test
    public void  testAutoCellWidthWrite() {
        sampleData.add(new WriteSampleData("定义数组\n" +
                "go中数组的定义方式如下。\n" +
                "\n" +
                "var 变量名 [长度]数组存储的类型\n" +
                "初始化数组\n" +
                "数组如果没有初始化，那么就是零值（比如int的零值是0，string的零值就是\"\"）。", 1, new Date()));
        EasyExcel.write("write.xlsx", WriteSampleData.class).sheet("模板").doWrite(sampleData);
    }
    @Test
    public void testTemplateBasicWrite() {
        WriteTemplateSampleData writeTemplateSampleData = new WriteTemplateSampleData();
        writeTemplateSampleData.setName("张三");
        writeTemplateSampleData.setAge(5);
        writeTemplateSampleData.setBirthday(new Date());
        EasyExcel.write("write.xlsx").withTemplate("template.xlsx").sheet().doFill(writeTemplateSampleData);
    }

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
}
