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
