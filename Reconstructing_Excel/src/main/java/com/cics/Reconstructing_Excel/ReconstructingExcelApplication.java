package com.cics.Reconstructing_Excel;

import org.springframework.boot.SpringApplication;
import java.io.FileInputStream;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;


@SpringBootApplication
public class ReconstructingExcelApplication {
	public static void main(String[] args) {
		Excel data = new Excel();
		data.writeExcel(data.readExcel());
	}
}
