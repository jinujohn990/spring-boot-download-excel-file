package com.jinu.excel.controller;

import java.io.ByteArrayInputStream;
import java.io.IOException;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.util.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import com.jinu.excel.service.ExportService;

@Controller
public class ExportController {
	@Autowired
	private ExportService exportService;
	@GetMapping("/download/employee.xlsx")
    public void downloadCsv(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=employee.xlsx");
        ByteArrayInputStream stream = exportService.employeetListToExcelFile();
        IOUtils.copy(stream, response.getOutputStream());
    }
}
