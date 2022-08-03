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
