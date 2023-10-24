package com.example.excel.feature.header;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelHeader {

    public final static Map<String, String> bookHeader = new HashMap();
    public final static Map<String, String> cdHeader = new HashMap();

    public final static List<String> bookHeaderForm = new ArrayList<>();
    public final static List<String> cdHeaderForm = new ArrayList<>();

    static{
        bookHeader.put("책번호", "bookId");
        bookHeader.put("책이름", "bookName");
        bookHeader.put("가격", "price");
        bookHeader.put("비고", "remark");

        cdHeader.put("CD번호", "CdId");
        cdHeader.put("CD이름", "CdName");
        cdHeader.put("가격", "price");
        cdHeader.put("비고", "remark");

        bookHeaderForm.add("책번호");
        bookHeaderForm.add("책이름");
        bookHeaderForm.add("가격");
        bookHeaderForm.add("비고");

        cdHeaderForm.add("CD번호");
        cdHeaderForm.add("CD이름");
        cdHeaderForm.add("가격");
        cdHeaderForm.add("비고");
    }
}
