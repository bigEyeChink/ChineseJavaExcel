package com.company;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;


public class Main {

    public static void main(String[] args) throws IOException {


        Workbook wb = new XSSFWorkbook();
        //CreationHelper 是一个用于往cell里面写入特殊数据的 类似 格式规范器 的一个东西
        //String 日期 超链接等一些复杂的东西 需要通过这样一个东西写入cell
        //类比   Scanner 		扫描 进 java
        // 		CreationHelper  写入 出 excel
        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");

        //Class Row 与sheet类似 没！有！ constructor 只能通过 Sheet 的instance creatRow来创建
        Row row = sheet.createRow(0);// createRow(行数); 行数从零开始! excel  文件从零开始 
        // Cell 举一反三 row
        Cell cell = row.createCell(0);


        // setCellValue() 函数用于cell内数据写入 parameter可以是int float boolean() 以及CreationHelper return回来的值
        cell.setCellValue(1); 

        // 写成一行为
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(
        		// createRichTextString() 将一个string 转化成 一个可以被 setCellValue接受的parameter
                createHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);// boolean会被转化成全部大写 TRUE

        // output
        try (OutputStream fileOut = new FileOutputStream("workbook.xls")) {
            wb.write(fileOut);
        }
    }
}
