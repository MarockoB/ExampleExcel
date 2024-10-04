package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

public class ExcelRead {
    public static ArrayList<Rectangle> readRectanglesFromExcel(String filePath) throws Exception {
        ArrayList<Rectangle> rectangles = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                double width = 0, height = 0;
                int cellIndex = 0;

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cellIndex) {
                        case 0:
                            width = (double) cell.getNumericCellValue();
                            break;
                        case 1:
                            height = (double) cell.getNumericCellValue();
                            break;
                    }
                    cellIndex++;
                }
                rectangles.add(new Rectangle(width, height));
            }
        }
        return rectangles;
    }

    public static void main(String[] args) throws Exception {
        ArrayList<Rectangle> rectangles = readRectanglesFromExcel("Example.xlsx");
        for (Object rect : rectangles) {
            System.out.println(rect);
        }
    }
}