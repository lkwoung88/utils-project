package com.feature.excel.object.impl;

import com.feature.excel.object.ExcelImporter;
import org.apache.poi.ss.usermodel.Sheet;

public class CdExcelImporter <T> extends ExcelImporter <T> {

    public CdExcelImporter(Class entityClass, Sheet sheet) {
        super(entityClass, sheet);
    }

    @Override
    public int saveEntities() {
        return 0;
    }
}
