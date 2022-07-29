package com.itlab1024.easyexcel.read;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.listener.PageReadListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.fastjson2.JSON;
import com.itlab1024.easyexcel.listener.GetHeadListener;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;

import java.util.List;

@Slf4j
public class EasyExcelReadTest {
    /**
     * 读取单个sheet。使用监听器，可以自定义也可以使用内置的PageReadListener
     */
    @Test
    public void testBasicRead() {
        EasyExcel.read("sample.xlsx", SampleData.class, new PageReadListener<SampleData>(dataList -> {
            log.info("读取到的数据是:{}", JSON.toJSONString(dataList));
        })).sheet().doRead();
    }
    @Test
    public void testMultiSheetsRead() {
        EasyExcel.read("sample.xlsx", SampleData.class, new PageReadListener<SampleData>(dataList -> {
            log.info("读取到的数据是:{}", JSON.toJSONString(dataList));
        })).doReadAll();
    }

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

    @Test
    public void testConvertRead() {
        EasyExcel.read("sample.xlsx", ConvertData.class, new PageReadListener<ConvertData>(dataList -> {
            log.info("读取到的数据是:{}", JSON.toJSONString(dataList));
        })).sheet().doRead();
    }

    /**
     * 同步读取
     */
    @Test
    public void testSyncRead() {
        List<Object> objects = EasyExcel.read("sample.xlsx").head(SampleData.class).sheet().doReadSync();
        log.info("读取结果{}", JSON.toJSONString(objects));
    }

    // 获取表头
    @Test
    public void testGetTableHeadRead() {
        EasyExcel.read("sample.xlsx", SampleData.class, new GetHeadListener()).sheet().doRead();
    }
}
