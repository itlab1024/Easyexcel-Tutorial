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
public class WriteSampleDataIndexed {
    @ExcelProperty(value = "姓名", index = 1)
    private String name;
    @ExcelProperty(value = "年龄", index = 2)
    private int age;
    @ExcelProperty(value = "出生年月", index = 0)
    @ColumnWidth(100)
    private Date birthday;
}
