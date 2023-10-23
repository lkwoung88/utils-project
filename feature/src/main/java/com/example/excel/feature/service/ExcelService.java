package com.example.excel.feature.service;

import com.example.excel.feature.entity.Book;
import com.example.excel.feature.entity.Cd;
import com.example.excel.feature.header.ExcelHeader;
import com.example.excel.feature.response.InputExcelResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Slf4j
@Service
public class ExcelService {

    public InputExcelResponse inputExcel(String kind, MultipartFile file) {

        InputExcelResponse response = new InputExcelResponse();

        try {
            log.info("IN TO ExcelService.excelToEntities");

            if ("book".equals(kind)) {
                this.ExcelToEntityList(Book.class, ExcelHeader.bookHeader, file);
            } else if ("cd".equals(kind)) {
                this.ExcelToEntityList(Cd.class, ExcelHeader.cdHeader, file);
            }

            return response;
        } catch (Exception e) {
            e.printStackTrace();
            return response;
        }
    }

    public <T> InputExcelResponse ExcelToEntityList(Class<T> entityClass, Map<String, String> header, MultipartFile file)
            throws IOException, NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException, NoSuchFieldException {

        log.info("IN TO ExcelService.ExcelToEntityList");

        InputExcelResponse response = new InputExcelResponse();
        List<T> entityList = new ArrayList<>();
        List<String> errorMessage = new ArrayList<>();
        boolean isSuccess = true;
        Workbook workbook = null;

        String extension = FilenameUtils.getExtension(file.getOriginalFilename());

        if ("xlsx".equals(extension)) {
            workbook = new XSSFWorkbook(file.getInputStream());
        } else if ("xls".equals(extension)) {
            workbook = new HSSFWorkbook(file.getInputStream());
        } else {
            System.out.println(" type error ");
            return null;
        }

        Sheet workSheet = workbook.getSheetAt(0);
        Row headerRow = workSheet.getRow(0);

        int totalRow = workSheet.getPhysicalNumberOfRows();
        int totalCol = headerRow.getPhysicalNumberOfCells();

        for (int rowIndex = 1; rowIndex < totalRow; rowIndex++) {
            Row row = workSheet.getRow(rowIndex);
            T entity = entityClass.getDeclaredConstructor().newInstance();

            for (int cellIndex = 0; cellIndex < totalCol; cellIndex++) {
                String cellValue = "";
                Cell cell = row.getCell(cellIndex);

                if (cell == null) {

                    if (isSuccess) {
                        isSuccess = false;
                    }

                    errorMessage.add(rowIndex + 1 + " row " + (char)('A' + cellIndex) +  " column");

                    continue;
                }

                CellType cellType = cell.getCellType();

                if ("NUMERIC".equals(cellType.toString())) {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                } else if ("STRING".equals(cellType.toString())) {
                    cellValue = cell.getStringCellValue();
                } else {
                    continue;
                }

                String headerKey = headerRow.getCell(cellIndex).getStringCellValue();
                String headerValue = header.get(headerKey);
                Field entityField = entityClass.getDeclaredField(headerValue);

                if (entityField != null) {
                    try {
                        entityField.setAccessible(true);

                        if (entityField.getType() == String.class) {
                            entityField.set(entity, cellValue);
                        } else if (entityField.getType() == short.class) {
                            entityField.set(entity, Short.parseShort(cellValue));
                        } else if (entityField.getType() == int.class) {
                            String[] split = cellValue.split("\\.");
                            entityField.set(entity, Integer.parseInt(split[0]));
                        } else if (entityField.getType() == long.class) {
                            entityField.set(entity, Long.parseLong(cellValue));
                        } else if (entityField.getType() == float.class) {
                            entityField.set(entity, Float.parseFloat(cellValue));
                        } else if (entityField.getType() == double.class) {
                            entityField.set(entity, Double.parseDouble(cellValue));
                        } else if (entityField.getType() == boolean.class) {
                            entityField.set(entity, Boolean.parseBoolean(cellValue));
                        } else if (entityField.getType() == char.class) {
                            entityField.set(entity, cellValue);
                        } else {

                        }

                    } catch (NumberFormatException e) {
                        isSuccess = false;
                        errorMessage.add(rowIndex + 1 + " row " + (char)('A' + cellIndex) +  " column");
                    } finally {
                        entityField.setAccessible(false);
                    }
                }
            }

            entityList.add(entity);
        }

        // test start
        for (T t : entityList) {
            System.out.println("t.toString() = " + t.toString());
        }

        for (String s : errorMessage) {
            System.out.println("s = " + s);
        }

        System.out.println("isSuccess = " + isSuccess);
        // test end

        response.setEntityList(entityList);
        response.setSuccess(isSuccess);
        response.setErrorMessage(errorMessage);

        return response;
    }
}
