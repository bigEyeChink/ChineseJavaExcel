package com.company;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;


public class Main {

    public static void main(String[] args) throws IOException {

    	// HSSF 是对老版本的excel文件定义 后缀xls H Horrible
        Workbook wb = new HSSFWorkbook();
        try  (OutputStream fileOut = new FileOutputStream("workbook.xls")) {
            wb.write(fileOut);
        }
        // XSSF 新版本 xlsx   X是后缀中多出来的x
        // X 和 H调用的jar不是完全相同的 有一部分相同 一部分不同 
        // 当不同部分未被加入lib时 报  java.lang.NoClassDefFoundError running time error
        Workbook wb2 = new XSSFWorkbook();

        try (OutputStream fileOut = new FileOutputStream("workbook.xlsx")) { //try 的意义在于 不知道workbook.xlsx 是否已经存在
            wb2.write(fileOut);
            //workbook instance函数  write（out流的变量名称） 功能:输出成品文件
        }

    }
}
