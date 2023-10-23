package com.example.excel.feature.response;

import lombok.Data;

import java.util.List;

@Data
public class InputExcelResponse<T> {

    private List<T> entityList;

    private boolean isSuccess;

    private List<String> errorMessage;
}
