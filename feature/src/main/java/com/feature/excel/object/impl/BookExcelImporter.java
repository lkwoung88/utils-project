package com.feature.excel.object.impl;

import com.feature.excel.object.ExcelImporter;
import org.apache.poi.ss.usermodel.Sheet;

public class BookExcelImporter<T> extends ExcelImporter <T> {

    public BookExcelImporter (Class<T> entityClass, Sheet sheet) {
        super(entityClass, sheet);
    }

    @Override
    public int saveEntities() {
        return 0;
    }
}
