package com.itlab1024.easyexcel.write;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class WriteTemplateSampleData {
    private String name;
    private int age;
    private Date birthday;
}
