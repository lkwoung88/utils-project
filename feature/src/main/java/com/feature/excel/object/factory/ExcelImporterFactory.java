package com.feature.excel.object.factory;

import com.feature.excel.entity.Book;
import com.feature.excel.entity.Cd;
import com.feature.excel.object.ExcelImporter;
import com.feature.excel.object.impl.BookExcelImporter;
import com.feature.excel.object.impl.CdExcelImporter;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.lang.Nullable;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelImporterFactory {

    @Nullable
    public static List<ExcelImporter> createExcelImporter(MultipartFile file) throws Exception {

        Workbook workbook = getWorkbook(file);

        if (workbook == null) {

            return null;
        }

        return createExcelImporter(workbook);
    }

    public static List<ExcelImporter> createExcelImporter(Workbook workbook) {

        int numberOfSheets = workbook.getNumberOfSheets();

        List<ExcelImporter> excelImporters = new ArrayList<>();

        for (int i = 0; i < numberOfSheets; i++) {

            Sheet sheet = workbook.getSheetAt(i);
            excelImporters.add(createExcelImporter(sheet));
        }

        return excelImporters;
    }

    @Nullable
    public static ExcelImporter createExcelImporter(Sheet sheet) {

        switch (sheet.getSheetName()) {

            case "book" -> {

                return new BookExcelImporter(Book.class, sheet);
            }
            case "cd" -> {

                return new CdExcelImporter(Cd.class, sheet);
            }
            default -> {

                return null;
            }
        }
    }

    private static Workbook getWorkbook(MultipartFile file) throws IOException {

        String extension = FilenameUtils.getExtension(file.getOriginalFilename());

        Workbook workbook = null;

        if ("xlsx".equals(extension)) {

            workbook = new XSSFWorkbook(file.getInputStream());
        }
        else if ("xls".equals(extension)) {

            workbook = new HSSFWorkbook(file.getInputStream());
        }
        else {

            System.out.println(" type error ");
            return null;
        }

        return workbook;
    }


}
