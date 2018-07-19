package com.company;

import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Calendar;
import java.util.Date;


public class Main {

    public static void main(String[] args) throws IOException {

        //创建新版本 excel
        Workbook wb = new XSSFWorkbook();
        // 创建写入器
        CreationHelper createHelper = wb.getCreationHelper();
        //创建新分页
        Sheet sheet = wb.createSheet("new sheet");
        //创建新分行 !!!INDEX 0开始
        Row row = sheet.createRow(0);
        //创建新格子 !!!INDEX 0开始
        Cell cell = row.createCell(0);

        //为什么 不用helper？？？？　设置成当前时间　但是这时候时间没有被style 是无法被看懂的数字
        cell.setCellValue(new Date());

        //在wb内创建一个cellstyle 来调整style of cell
        CellStyle cellStyle = wb.createCellStyle();
        //更改dataFormat  setDataFormat(arg是一个short)
        // 这行代码中有createHelper.createDataFormat().getFormat("m/d/yy h:mm") 产生
        cellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
        //将cell的指针指向index为1 在row 中的 entry   可以将 row看做 cell 的sequence
        cell = row.createCell(1);
        // 先设置value, 再设置style
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);

        // 相对简单的 可以直接使用calender class
        // !!! 如果不用setstyle calender 和date 所return 的value 相同
        cell = row.createCell(2);
        cell.setCellValue(Calendar.getInstance());
        // 仍然先设置value
        cell.setCellStyle(cellStyle);
        try (OutputStream fileOut = new FileOutputStream("workbook.xls")) {
            wb.write(fileOut);
        }
    }
}
