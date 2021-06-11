package com.seleniumframework.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Xcel {
//public static void main(String[] args) throws Exception {
	public static String a;
public static  void main() throws Exception {
	//File src =new File("./xcel1/a.xlsx");
	File src =new File("./xcel1/a.xlsx");
	FileInputStream fil =new FileInputStream(src);
	XSSFWorkbook fi=new XSSFWorkbook(fil);
	XSSFSheet al= fi.getSheetAt(1);
    a=al.getRow(1).getCell(2).getStringCellValue();
    System.out.println(a);
    fi.close();
}
}
