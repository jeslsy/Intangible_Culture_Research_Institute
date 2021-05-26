package com.cics.Reconstructing_Excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

public class test {
    // 엑셀에 내용 쓰기
    public void testExcel(ArrayList<exData> list) {
        try {
            Workbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet("My Sample Excel");


















            //너비를 n 문자 너비로 설정 = 문자 수 * 256
            //int widthUnits = 20*256;
            //sheet.setColumnWidth(1, widthUnits);

            //높이를 twips로 n 포인트로 설정 = n * 20
            //short heightUnits = 60*20;
            //cell.getRow().setHeight(heightUnits);

            //Excel 파일 작성
            FileOutputStream fileOut = null;
            fileOut = new FileOutputStream("myFile.xlsx");
            wb.write(fileOut);
            fileOut.close();

        } catch (IOException ioex) {
        }
    }

}
