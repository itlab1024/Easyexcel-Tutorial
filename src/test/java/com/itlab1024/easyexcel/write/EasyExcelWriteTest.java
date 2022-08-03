package com.itlab1024.easyexcel.write;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.data.ImageData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.util.FileUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
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
        EasyExcel.write("write.xlsx", WriteImageSampleData.class).sheet().doWrite(Collections.singleton(imageSampleData));
    }
}
