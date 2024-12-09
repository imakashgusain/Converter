package com.converter.excelJson.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelUtils {

    public static List<List<String>> readExcelFile(File file) throws IOException {
        List<List<String>> data = new ArrayList<>();
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0); // Reading the first sheet

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            List<String> rowData = new ArrayList<>();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                rowData.add(cell.toString()); // You can add specific formatting here
            }
            data.add(rowData);
        }
        workbook.close();
        return data;
    }

    public static File createExcelFileFromJson(List<List<String>> jsonData, String fileName) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        int rowNum = 0;
        for (List<String> rowData : jsonData) {
            Row row = sheet.createRow(rowNum++);
            int cellNum = 0;
            for (String cellData : rowData) {
                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(cellData);
            }
        }

        // Save to a temporary file
        File tempFile = File.createTempFile(fileName, ".xlsx");
        FileOutputStream fileOut = new FileOutputStream(tempFile);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();

        return tempFile;
    }

}
