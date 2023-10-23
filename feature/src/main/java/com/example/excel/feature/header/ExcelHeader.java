package com.example.excel.feature.header;

import java.util.HashMap;
import java.util.Map;

public class ExcelHeader {

    public final static Map<String, String> bookHeader = new HashMap();
    public final static Map<String, String> cdHeader = new HashMap();

    static{
        bookHeader.put("책번호", "bookId");
        bookHeader.put("책이름", "bookName");
        bookHeader.put("가격", "price");
        bookHeader.put("비고", "remark");

        cdHeader.put("CD번호", "CdId");
        cdHeader.put("CD이름", "CdName");
        cdHeader.put("가격", "price");
        cdHeader.put("비고", "remark");
    }
}
