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
