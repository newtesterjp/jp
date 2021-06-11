package com.seleniumframework.excel;

import java.io.File;


import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writeexcelapachepoi {
	public static void main(String[] args) throws IOException {
		File src =new File("C:\\Users\\jp\\OneDrive\\Desktop\\pract1.xlsx");
		FileInputStream fil =new FileInputStream(src);
		XSSFWorkbook fi=new XSSFWorkbook(fil);
		XSSFSheet al= fi.getSheetAt(1);
		al.getRow(1).createCell(1).setCellValue("hihello");
		FileOutputStream fout=new FileOutputStream(src);
		fi.write(fout);
		fi.close();
	}

}
