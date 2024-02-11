package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

public class JsonToExcelConverter {
	public static void main(String[] args) {
        try {
            // Read JSON data from a file (replace with your JSON file)
            ObjectMapper objectMapper = new ObjectMapper();
            JsonNode jsonData = objectMapper.readTree(new File("C:\\Users\\sukee\\OneDrive\\Desktop\\ExceltoJsonConvertor\\sample.json"));

            // Create Excel workbook and sheet
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sheet1");

            // Write JSON data to Excel
            int rowNum = 0;
            for (JsonNode jsonRow : jsonData) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;

                for (JsonNode jsonCell : jsonRow) {
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(jsonCell.asText());
                }
            }

            // Save the workbook to a file (replace with your desired output path)
            try (FileOutputStream fileOut = new FileOutputStream("C:\\Users\\sukee\\OneDrive\\Desktop\\ExceltoJsonConvertor\\JsontoExcel.xlsx")) {
                workbook.write(fileOut);
            }

            // Close workbook
            workbook.close();

            System.out.println("Conversion completed successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
