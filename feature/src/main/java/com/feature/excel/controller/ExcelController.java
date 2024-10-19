package com.feature.excel.controller;

import com.feature.excel.response.InputExcelResponse;
import com.feature.excel.service.ExcelService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.formula.functions.T;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

@RestController
@RequiredArgsConstructor
@RequestMapping("/feature/excel")
public class ExcelController {

    private final ExcelService excelService;

    @PostMapping("/import")
    public ResponseEntity<T> importExcel(@RequestParam("file") MultipartFile file) throws Exception {

        return excelService.importExcel(file);
    }

    // Excel To Entity List
    @PostMapping("/input")
    public InputExcelResponse inputExcel(@RequestPart(value = "kind") String kind, @RequestParam("excel") MultipartFile file) {
        return excelService.inputExcel(kind, file);
    }

    // Download Excel Form
    @GetMapping("/form/{kind}")
    public void downloadExcelForm(@PathVariable String kind, HttpServletResponse response) {
        excelService.downloadExcelForm(kind, response);
    }
}
