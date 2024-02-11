package com.example;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.ObjectMapper;

public class ExcelToJsonConverter1 {
	public static void main(String[] args) {
        // Read sheet name from console input
        Scanner scanner = new Scanner(System.in);
        System.out.print("Enter the sheet name: ");
        String sheetNameToPrint = scanner.nextLine();

        try {
            // Load Excel file
            FileInputStream file = new FileInputStream("C:\\Users\\sukee\\OneDrive\\Desktop\\ExceltoJsonConvertor\\ExceltoJson.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Print data for the specified sheet
            printSheetData(workbook, sheetNameToPrint);

            // Close workbook
            workbook.close();
            file.close();

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // Close the scanner
            scanner.close();
        }
    }

    private static void printSheetData(XSSFWorkbook workbook, String sheetName) {
        // Get the specified sheet
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(sheetName);

        // Check if the sheet exists
        if (sheet != null) {
            // Iterate through rows and columns
            Iterator<Row> rowIterator = sheet.iterator();

            // create json array Print data
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                // Print data for each row. create json object for each row
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    System.out.print(cell.toString() + " ");
                }
                System.out.println();
            }
        } else {
            System.out.println("Sheet '" + sheetName + "' not found.");
        }
    }
}
