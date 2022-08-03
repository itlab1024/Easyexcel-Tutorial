package com.itlab1024.easyexcel.write;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.*;
import com.alibaba.excel.enums.BooleanEnum;
import com.alibaba.excel.enums.poi.HorizontalAlignmentEnum;
import com.alibaba.excel.enums.poi.VerticalAlignmentEnum;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Font;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
@HeadRowHeight(30) // 表头行高
@ContentRowHeight(50) // 内容行高
@ColumnWidth(30) // 列宽
@ContentFontStyle(fontName = "monaco", bold = BooleanEnum.TRUE, color = Font.COLOR_RED, underline = Font.U_DOUBLE) // 内容文字风格
@HeadFontStyle(fontName = "Arial", bold = BooleanEnum.TRUE, color = Font.COLOR_RED, underline = Font.U_SINGLE_ACCOUNTING) // 表头文字风格
@HeadStyle(horizontalAlignment = HorizontalAlignmentEnum.LEFT, verticalAlignment = VerticalAlignmentEnum.CENTER) //表头风格
@OnceAbsoluteMerge(firstRowIndex = 5, lastRowIndex = 6, firstColumnIndex = 1, lastColumnIndex = 2)
public class WriteSampleDataAnnotation {
    @ExcelProperty("姓名")
    @ContentLoopMerge(eachRow = 2)
    private String name;
    @ExcelProperty("年龄")
    private int age;
    @ExcelProperty("出生年月")
    @ColumnWidth(50) // 单独设置 birthday列宽
    private Date birthday;
}
