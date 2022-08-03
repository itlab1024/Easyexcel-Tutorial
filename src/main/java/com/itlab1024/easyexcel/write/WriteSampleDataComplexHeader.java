package com.itlab1024.easyexcel.write;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class WriteSampleDataComplexHeader {
    @ExcelProperty({"基本信息", "姓名"})
    private String name;
    @ExcelProperty(value = {"基本信息", "年龄"})
    private int age;
    @ExcelProperty(value = "出生年月")
    @ColumnWidth(30)
    private Date birthday;
}
