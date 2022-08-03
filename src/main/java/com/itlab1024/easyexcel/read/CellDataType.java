package com.itlab1024.easyexcel.read;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.data.CellData;
import lombok.Data;

@Data
public class CellDataType {
    @ExcelProperty("公式")
    private CellData<String> formula;
}
