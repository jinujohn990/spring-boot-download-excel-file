package com.jinu.excel.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

import com.jinu.excel.model.Employee;
import com.jinu.excel.repository.EmployeeRepository;



@Service
public class ExportService {
	@Autowired
	private EmployeeRepository employeeRepository;
	
	public  ByteArrayInputStream employeetListToExcelFile() {
		List<Employee> employeeList = null;
		try(Workbook workbook = new XSSFWorkbook()){
			Sheet sheet = workbook.createSheet("Customers");
			
			Row row = sheet.createRow(0);
	        CellStyle headerCellStyle = workbook.createCellStyle();
	        headerCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
	        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        // Creating header
	        Cell cell = row.createCell(0);
	        cell.setCellValue("emp_id");
	        cell.setCellStyle(headerCellStyle);
	        
	        cell = row.createCell(1);
	        cell.setCellValue("First name");
	        cell.setCellStyle(headerCellStyle);
	
	        cell = row.createCell(2);
	        cell.setCellValue("Last name");
	        cell.setCellStyle(headerCellStyle);
	
	        cell = row.createCell(3);
	        cell.setCellValue("Email");
	        cell.setCellStyle(headerCellStyle);
	        sheet.autoSizeColumn(0);
	        sheet.autoSizeColumn(1);
	        sheet.autoSizeColumn(2);
	        sheet.autoSizeColumn(3);
            int offset =0;
            int pageNum = 0;
            long employeeCount = employeeRepository.count();
            while(offset<employeeCount) {
            	int lastRow =  sheet.getLastRowNum();
            	employeeList = pagedEmployeeSearchResult(pageNum);
            	 for(int i = 0; i < employeeList.size(); i++) {
     	        	Row dataRow = sheet.createRow(lastRow + i+1);
     	        	dataRow.createCell(0).setCellValue(employeeList.get(i).getEmpId());
     	        	dataRow.createCell(1).setCellValue(employeeList.get(i).getFirstName());
     	        	dataRow.createCell(2).setCellValue(employeeList.get(i).getLastName());
     	        	dataRow.createCell(3).setCellValue(employeeList.get(i).getEmail());
     	        }
            	offset= offset+10;
        	pageNum++;
            }
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
	        workbook.write(outputStream);
	        return new ByteArrayInputStream(outputStream.toByteArray());	
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return null;
		
	}

	private List<Employee> pagedEmployeeSearchResult(int offset) {
		Pageable paging = PageRequest.of(offset, 10, Sort.by("empId"));
		 Page<Employee> pagedResult = employeeRepository.findAll(paging);
		 if(pagedResult.hasContent()) {
	            return pagedResult.getContent();
	        } else {
	            return new ArrayList<Employee>();
	        }
	}
	
}
