package com.feature.excel.object;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public abstract class ExcelImporter <T> {

    Sheet sheet;
    List<T> entities = new ArrayList<>();
    Class<T> clazz;

    public ExcelImporter(Class<T> clazz, Sheet sheet) {

        this.clazz = clazz;
        this.sheet = sheet;
    }

    public boolean importExcel(){

        try {

            this.transferExcelToEntities();
            this.saveEntities();
            return true;
        }
        catch (Exception e) {

            log.error(e.getMessage());
            return false;
        }
    }

    protected void transferExcelToEntities() throws Exception {

        Row headerRow = this.sheet.getRow(0);

        int rowNum = this.sheet.getPhysicalNumberOfRows();
        int columnNum = headerRow.getPhysicalNumberOfCells();

        for (int rowIndex = 1; rowIndex < rowNum; rowIndex++) {

            Row valueRow = this.sheet.getRow(rowIndex);

            T entity = this.clazz.getDeclaredConstructor().newInstance();

            setFieldValue(columnNum, valueRow, headerRow, entity);

            this.entities.add(entity);
        }

    }

    private void setFieldValue(int columnNum, Row valueRow, Row headerRow, T entity) throws IllegalAccessException {

        for (int cellIndex = 0; cellIndex < columnNum; cellIndex++) {

            String cellValue = "";
            Cell cell = valueRow.getCell(cellIndex);

            if (cell == null) {
                continue;
            }

            CellType cellType = cell.getCellType();

            if ("NUMERIC".equals(cellType.toString())) {

                cellValue = String.valueOf((int)cell.getNumericCellValue());
            } else if ("STRING".equals(cellType.toString())) {

                cellValue = cell.getStringCellValue();
            } else {

                continue;
            }

            String headerName = headerRow.getCell(cellIndex).getStringCellValue();
            Field entityField = null;

            Field[] fields = clazz.getDeclaredFields();

            for (Field field : fields) {

                if (field.isAnnotationPresent(ExcelHeader.class)) {

                    ExcelHeader excelHeader = field.getAnnotation(ExcelHeader.class);

                    if (excelHeader.name().equals(headerName)) {

                        entityField = field;
                    }
                }
            }

            if (entityField != null) {
                try {
                    entityField.setAccessible(true);

                    if (entityField.getType() == String.class) {

                        entityField.set(entity, cellValue);
                    }
                    else if (entityField.getType() == short.class) {

                        entityField.set(entity, Short.parseShort(cellValue));
                    }
                    else if (entityField.getType() == int.class) {

                        String[] split = cellValue.split("\\.");
                        entityField.set(entity, Integer.parseInt(split[0]));
                    }
                    else if (entityField.getType() == long.class) {

                        entityField.set(entity, Long.parseLong(cellValue));
                    }
                    else if (entityField.getType() == float.class) {

                        entityField.set(entity, Float.parseFloat(cellValue));
                    }
                    else if (entityField.getType() == double.class) {

                        entityField.set(entity, Double.parseDouble(cellValue));
                    }
                    else if (entityField.getType() == boolean.class) {

                        entityField.set(entity, Boolean.parseBoolean(cellValue));
                    }
                    else if (entityField.getType() == char.class) {

                        entityField.set(entity, cellValue);
                    }
                    else {

                    }
                }
                catch (NumberFormatException e) {

                    log.error(e.getMessage());
                }
                finally {

                    entityField.setAccessible(false);
                }
            }
        }
    }

    protected abstract int saveEntities();
}
