package com.company;

import java.io.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
        public static void main(String[] args)throws Exception {
        	
            Workbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet("new sheet");

            // Create a row and put some cells in it. Rows are 0 based.
            Row row = sheet.createRow(1);

            // Create a cell and put a value in it.
            Cell cell = row.createCell(1);
            cell.setCellValue(4);

            // Style the cell with borders all around.
            CellStyle style = wb.createCellStyle();
            style.setBorderBottom(BorderStyle.THIN);
            style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            style.setBorderLeft(BorderStyle.THIN);
            style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
            style.setBorderRight(BorderStyle.THIN);
            style.setRightBorderColor(IndexedColors.BLUE.getIndex());
            style.setBorderTop(BorderStyle.MEDIUM_DASHED);
            style.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cell.setCellStyle(style);

            // Write the output to a file
            try (OutputStream fileOut = new FileOutputStream("boders.xlsx")) {
                wb.write(fileOut);
            }

            wb.close();

        }
    }
