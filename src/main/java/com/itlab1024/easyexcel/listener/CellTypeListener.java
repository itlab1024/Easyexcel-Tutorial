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
