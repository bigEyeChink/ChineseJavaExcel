package com.company;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;


public class Main {

    public static void main(String[] args) throws IOException {


        Workbook wb = new XSSFWorkbook();
        // Sheet是表格中的一页  Sheet没！有！constructor 只能通过 WorkBook 的instance creatSheet来创建
        Sheet sheet1 = wb.createSheet("new sheet");//createSheet(String sheet的名字)
        Sheet sheet2 = wb.createSheet("second sheet");


        // sheet 命名注意事项
        // Note that sheet name is Excel must not exceed 31 characters
        // and must not contain any of the any of the following characters:
        // 0x0000
        // 0x0003
        // colon (:)
        // backslash (\)
        // asterisk (*)
        // question mark (?)
        // forward slash (/)
        // opening square bracket ([)
        // closing square bracket (])

        // WorkbookUtil.createSafeSheetName(String 想要使用的名字) 可以return 一个把不合法的字符替！换！成空！格！的名字 
      
        String safeName = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]"); // returns " O'Brien's sales   "
        Sheet sheet3 = wb.createSheet(safeName);

        // XSSFWorkbook被写成xls文件是，程序执行 当用excel打开文件时 高版本excel会开启兼容模式读取文件
        // 无论是什么的WorkBook被写成xlsx文件后 没有办法在低版本的excel文件中打开
        try (OutputStream fileOut = new FileOutputStream("workbook.xls")) {
            wb.write(fileOut);
        }
    }
}
