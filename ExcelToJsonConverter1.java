package com.example;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.ObjectMapper;

public class ExcelToJsonConverter1 {
	public static void main(String[] args) {
        try {
            // Load Excel file
            FileInputStream file = new FileInputStream("C:\\Users\\sukee\\OneDrive\\Desktop\\ExceltoJsonConvertor\\ExceltoJson.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Assuming the data is in the first sheet
            Iterator<Row> rowIterator = workbook.getSheetAt(0).iterator();

            // Create ObjectMapper for JSON conversion
            ObjectMapper objectMapper = new ObjectMapper();

            // Create JSON array
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                // Create JSON object for each row
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    System.out.print(cell.toString() + " ");
                }
                System.out.println();
            }

            // Close workbook
            workbook.close();
            file.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
