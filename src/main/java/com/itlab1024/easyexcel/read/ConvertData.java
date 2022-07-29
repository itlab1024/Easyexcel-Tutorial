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
    private String birthday;
}
