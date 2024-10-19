package com.feature.excel.entity;

import com.feature.excel.object.ExcelHeader;
import lombok.Data;

@Data
public class Book {

    @ExcelHeader(name = "일련번호")
    private String uid;

    @ExcelHeader(name = "분류코드")
    private String classificationCode;

    @ExcelHeader(name = "이름")
    private String name;

    @ExcelHeader(name = "가격")
    private int price;

    @ExcelHeader(name = "비고")
    private String remark;
}
