package com.company;

import java.io.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
        public static void main(String[] args)throws Exception {
        	//创建新的wookbook xlxs
            Workbook wb = new XSSFWorkbook(); //or new HSSFWorkbook();
            //创建新的sheet页面
            Sheet sheet = wb.createSheet();
            //创建新的行 在第三行的位置
            Row row = sheet.createRow(2);
            //讲整个第三行的高度设置为30个单位
            row.setHeightInPoints(30);
            //利用函数更改cell的内容
            //参数解释
            //水平方向对齐和垂直方向对齐应用的参数 都是Final int Constant 用枚举法列举出的情况 底层实际上是数字
            //起名字是为了方便使用
            //GENERAL 为默认  text左对齐 boolean居中 数字时间日期右对齐
            //LEFT 改为左对齐 RIGHT CENTER 同理
            //FULL 单元格调整为内容合适的宽度 
            //JUSTIFY 内容调整为适合单元格宽度

            createCell(wb, row, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM);
            createCell(wb, row, 1, HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.BOTTOM);
            createCell(wb, row, 2, HorizontalAlignment.FILL, VerticalAlignment.CENTER);
            createCell(wb, row, 3, HorizontalAlignment.GENERAL, VerticalAlignment.CENTER);
            createCell(wb, row, 4, HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY);
            createCell(wb, row, 5, HorizontalAlignment.LEFT, VerticalAlignment.TOP);
            createCell(wb, row, 6, HorizontalAlignment.RIGHT, VerticalAlignment.TOP);

            // Write the output to a file
            try (OutputStream fileOut = new FileOutputStream("xssf-align.xlsx")) {
                wb.write(fileOut);
            }

            wb.close();
        }

        /**
         * Creates a cell and aligns it a certain way.
         *
         * @param wb     the workbook
         * @param row    the row to create the cell in
         * @param column the column number to create the cell in
         * @param halign the horizontal alignment for the cell.
         * @param valign the vertical alignment for the cell.
         */
        private static void createCell(Workbook wb, Row row, int column, HorizontalAlignment halign, VerticalAlignment valign) {
            Cell cell = row.createCell(column);
            cell.setCellValue("Align It");
            //创建一个cellstyle 来规定一种格式
            CellStyle cellStyle = wb.createCellStyle();
            //setAlignment 可以调整水平方向上的对齐
            cellStyle.setAlignment(halign);
            //setVerticalAlignment 调整垂直方向上的对齐
            cellStyle.setVerticalAlignment(valign);
            //将style应用到cell上
            cell.setCellStyle(cellStyle);
        }
    }
