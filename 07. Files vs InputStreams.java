package com.company;

import java.io.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

public class Main {

    public static void main(String[] args) throws IOException, InvalidFormatException {
    	//读取现有的excel文件 两大类选择
    	//第一种Workbookfactory
    	//是poi自带的直接读取进workbook 不用区分xls xlsx的一种方法
    	//可以直接用factory直接打开一个文件地址
        Workbook wb = WorkbookFactory.create(new File("MyExcel.xlsx"));

        //也可以使用inputStrem InputStream 在读取时需要更多的内存 但可以提供buffer
        // 问题来了  buffer 的作用是什么？
        //Workbook wbStream = WorkbookFactory.create(new FileInputStream("MyExcel2.xlsx"));


        // Write the output to a file
        try (OutputStream fileOut = new FileOutputStream("MyExcelCopy.xlsx")) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }




        //另外一种方法是使用NPOIFSFileSystem class
        //暂时没有使用过，但使用应该与Scanner作用类似  可以部分读取 
        // 在结尾时也需要关闭 NPOIFSFileSystem 
		 // HSSFWorkbook, File
		  NPOIFSFileSystem fs = new NPOIFSFileSystem(new File("file.xls"));
		  HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
		  ....
		  fs.close();

		  // HSSFWorkbook, InputStream, needs more memory
		  NPOIFSFileSystem fs = new NPOIFSFileSystem(myInputStream);
		  HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);



		  // 与xls文件不同，新版excel文件需要用OPCPackage 读取， 作用相似
		  // XSSFWorkbook, File
		  OPCPackage pkg = OPCPackage.open(new File("file.xlsx"));
		  XSSFWorkbook wb = new XSSFWorkbook(pkg);
		  ....
		  pkg.close();

		  // XSSFWorkbook, InputStream, needs more memory
		  OPCPackage pkg = OPCPackage.open(myInputStream);
		  XSSFWorkbook wb = new XSSFWorkbook(pkg);
		  ....
		  pkg.close();       



    }
}
