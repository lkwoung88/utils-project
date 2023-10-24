package com.example.excel.feature.controller;

import com.example.excel.feature.response.InputExcelResponse;
import com.example.excel.feature.service.ExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

@RestController
@RequiredArgsConstructor
@RequestMapping("/feature/excel")
public class ExcelController {

    private final ExcelService excelService;

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
