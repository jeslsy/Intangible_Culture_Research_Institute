package com.cics.Reconstructing_Excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

public class save_middle {

    // 엑셀 파일 읽기 완료 코드
    public ArrayList<exData> dExcel() {
        ArrayList<exData> list = new ArrayList<exData>();

        try {
            FileInputStream file = new FileInputStream("C:/Users/CICS/OneDrive/바탕 화면/성남엑셀매크로/test_directory/test_read.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            int rowindex=0;
            int columnindex=0;

            //시트 수 (첫번째에만 존재하므로 0을 준다)
            // 만약 여러 시트를 읽기위해서는 FOR문을 한번더 돌려준다
            XSSFSheet sheet=workbook.getSheetAt(0);

            // 행의 수
            int rows=sheet.getPhysicalNumberOfRows();
            for(rowindex=2;rowindex<rows;rowindex++){

                exData ed = new exData();

                //행을읽는다
                XSSFRow row=sheet.getRow(rowindex);

                if(row !=null){

                    //셀의 수
                    int cells=row.getPhysicalNumberOfCells();
                    for(columnindex=0; columnindex<=cells; columnindex++){

                        //셀값을 읽는다
                        XSSFCell cell=row.getCell(columnindex);
                        String value="";

                        if(columnindex == 3 || columnindex == 4 || columnindex == 9 || columnindex == 10){
                            //셀이 빈값일경우를 위한 널체크
                            if(cell==null){
                                continue; }
                            else {
                                //타입별로 내용 읽기
                                switch (cell.getCellType()){
                                    case XSSFCell.CELL_TYPE_FORMULA:
                                        value=cell.getCellFormula();
                                        break;
                                    case XSSFCell.CELL_TYPE_NUMERIC:
                                        value=cell.getNumericCellValue()+"";
                                        break;
                                    case XSSFCell.CELL_TYPE_STRING:
                                        value=cell.getStringCellValue()+"";
                                        break;
                                    case XSSFCell.CELL_TYPE_BLANK:
                                        value=cell.getBooleanCellValue()+"";
                                        break;
                                    case XSSFCell.CELL_TYPE_ERROR:
                                        value=cell.getErrorCellValue()+"";
                                        break; }
                            }
                            if(columnindex == 3)
                                ed.setData_num(value);
                            else if(columnindex == 4)
                                ed.setTitle(value);
                            else if(columnindex == 9)
                                ed.setContents(value);
                            else if(columnindex == 10)
                                ed.setAddress(value);
                            System.out.println(value);
                        }
                    }
                }
                // list에 넣기
                list.add(ed);
            }
        }catch(Exception e) { e.printStackTrace(); }
        System.out.println("-------------------------------------------------");
        return list;
    }

    // 엑셀 파일 읽기
    public ArrayList<exData> asf() {
        ArrayList<exData> list = new ArrayList<exData>();

        try {
            FileInputStream file = new FileInputStream("C:/Users/CICS/OneDrive/바탕 화면/성남엑셀매크로/test_directory/조각test.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            int rowindex=0;
            int columnindex=0;

            //시트 수 (첫번째에만 존재하므로 0을 준다)
            // 만약 각 시트를 읽기위해서는 FOR문을 한번더 돌려준다
            XSSFSheet sheet=workbook.getSheetAt(0);

            // 행의 수
            int rows=sheet.getPhysicalNumberOfRows();
            for(rowindex=2;rowindex<rows;rowindex++){

                exData ed = new exData();

                //행을읽는다
                XSSFRow row=sheet.getRow(rowindex);

                if(row !=null){

                    //셀의 수
                    int cells=row.getPhysicalNumberOfCells();
                    for(columnindex=0; columnindex<=cells; columnindex++){

                        //셀값을 읽는다
                        XSSFCell cell=row.getCell(columnindex);
                        String value="";

                        if(columnindex == 3 || columnindex == 4 || columnindex == 9 || columnindex == 10){
                            //셀이 빈값일경우를 위한 널체크
                            if(cell==null){
                                continue; }
                            else {
                                //타입별로 내용 읽기
                                switch (cell.getCellType()){
                                    case XSSFCell.CELL_TYPE_FORMULA:
                                        value=cell.getCellFormula();
                                        break;
                                    case XSSFCell.CELL_TYPE_NUMERIC:
                                        value=cell.getNumericCellValue()+"";
                                        break;
                                    case XSSFCell.CELL_TYPE_STRING:
                                        value=cell.getStringCellValue()+"";
                                        break;
                                    case XSSFCell.CELL_TYPE_BLANK:
                                        value=cell.getBooleanCellValue()+"";
                                        break;
                                    case XSSFCell.CELL_TYPE_ERROR:
                                        value=cell.getErrorCellValue()+"";
                                        break; }
                            }
                            if(columnindex == 3)
                                ed.setData_num(value);
                            else if(columnindex == 4)
                                ed.setTitle(value);
                            else if(columnindex == 9)
                                ed.setContents(value);
                            else if(columnindex == 10)
                                ed.setAddress(value);
                            System.out.println(value);
                        }
                    }
                }
                // list에 넣기
                list.add(ed);
            }
            System.out.println(list.toString());
        }catch(Exception e) { e.printStackTrace(); }
        return list;
    }


    // 엑셀에 내용 쓰기 실패 코드
    public void writeExcel(ArrayList<exData> list) {
        try{
            // 엑셀 파일 열기
            FileOutputStream file = new FileOutputStream("C:/Users/CICS/OneDrive/바탕 화면/성남엑셀매크로/test_directory/조각test.xlsx");
            XSSFWorkbook xworkbook = new XSSFWorkbook();

            // 시트 생성
            XSSFSheet xsheet = xworkbook.createSheet("정리");

            // 이미지 넣을 셀 병합 (1번 3번)
            for(int i=0;i<66;i+=3){
                xsheet.addMergedRegion(new CellRangeAddress(
                        i,
                        i+1,
                        1,
                        1
                ));
                xsheet.addMergedRegion(new CellRangeAddress(
                        i,
                        i+1,
                        4,
                        4
                ));
            }

            // 각 행 읽기
            XSSFRow curRow;

            //리스트 크기
            int row = list.size();
            Cell cell = null;


            // 시트 꾸미기


            int cnt = 0;
            // 시트 내용
            for(int i=0; i<66; i++){

                //row 생성
                curRow = xsheet.createRow(i);

                cell = curRow.createCell(2);
                cell.setCellValue(String.valueOf(list.get(cnt).getTitle()));
                cell = curRow.createCell(5);
                cell.setCellValue(String.valueOf(list.get(cnt+1).getTitle()));

                //row 생성
                curRow = xsheet.createRow(++i);

                cell = curRow.createCell(2);
                cell.setCellValue(String.valueOf(list.get(cnt).getContents()));
                cell = curRow.createCell(5);
                cell.setCellValue(String.valueOf(list.get(cnt+1).getContents()));

                //row 생성
                curRow = xsheet.createRow(++i);

                cell = curRow.createCell(1);
                cell.setCellValue(String.valueOf(list.get(cnt).getData_num()));
                cell = curRow.createCell(2);
                cell.setCellValue(String.valueOf(list.get(cnt+1).getAddress()));
                cell = curRow.createCell(4);
                cell.setCellValue(String.valueOf(list.get(cnt).getData_num()));
                cell = curRow.createCell(5);
                cell.setCellValue(String.valueOf(list.get(cnt+1).getAddress()));

                cnt+=2;
            }

            //열 너비 설정
            for(int i=1;i<6;i++) {
                xsheet.autoSizeColumn(i);
                xsheet.setColumnWidth(i, (xsheet.getColumnWidth(i))+256);
            }

            xworkbook.write(file);
            file.close();

        } catch(Exception e) {
            e.printStackTrace();
        }
    }

    // 셀 병합 완료 코드
    public void writExcel(ArrayList<exData> list) {
        try{
            // 엑셀 파일 열기
            FileOutputStream file = new FileOutputStream("C:/Users/CICS/OneDrive/바탕 화면/성남엑셀매크로/test_directory/test_write.xlsx");
            XSSFWorkbook xworkbook = new XSSFWorkbook();

            // 시트 생성
            XSSFSheet xsheet = xworkbook.createSheet("정리");

            // 리스트 크기
            int list_row = list.size();


            // 이미지 넣을 셀 병합 (1번 3번)
            for(int i=0;i<list_row*4/2;i+=4){
                xsheet.addMergedRegion(new CellRangeAddress(
                        i,
                        i+1,
                        0,
                        0
                ));
                xsheet.addMergedRegion(new CellRangeAddress(
                        i,
                        i+1,
                        3,
                        3
                ));
            }

            // 시트 꾸미기

            // 시트 내용


            xworkbook.write(file);
            file.close();

        } catch(Exception e) {
            e.printStackTrace();
        }
    }

    // 서식 아직 내용 완성.
    public void write(ArrayList<exData> list) {
        try{
            // 엑셀 파일 열기
            FileOutputStream file = new FileOutputStream("C:/Users/CICS/OneDrive/바탕 화면/성남엑셀매크로/test_directory/test_write.xlsx");
            XSSFWorkbook xworkbook = new XSSFWorkbook();

            // 시트 생성
            XSSFSheet xsheet = xworkbook.createSheet("정리");
            XSSFRow curRow;

            // 리스트 크기
            int row = list.size();

            Cell cell = null;

            // 리스트 문제 없이 잘 넘어옴.
            for(int i=0;i<row;i++){
                System.out.println(list.get(i).toString());
            }

            // list 증가
            int list_cnt = 0;

            // 시트 내용(list 순서)
            for(int cur_row = 0; cur_row<row*4/2; cur_row++){   // 행 개수
                // 행 생성
                curRow = xsheet.createRow(cur_row);
                if(cur_row % 4 == 0){
                    cell = curRow.createCell(1);
                    cell.setCellValue(list.get(list_cnt).getTitle());

                    cell = curRow.createCell(4);
                    cell.setCellValue(list.get(list_cnt+1).getTitle());
                }
                else if(cur_row % 4 == 1){
                    cell = curRow.createCell(1);
                    cell.setCellValue(list.get(list_cnt).getContents());

                    cell = curRow.createCell(4);
                    cell.setCellValue(list.get(list_cnt + 1).getContents());
                }
                else if(cur_row % 4 == 2){
                    cell = curRow.createCell(0);
                    cell.setCellValue(list.get(list_cnt).getData_num());

                    cell = curRow.createCell(1);
                    cell.setCellValue(list.get(list_cnt).getAddress());

                    cell = curRow.createCell(3);
                    cell.setCellValue(list.get(list_cnt + 1).getData_num());

                    cell = curRow.createCell(4);
                    cell.setCellValue(list.get(list_cnt + 1).getAddress());

                    list_cnt+=2;
                }

            }

            //열 너비 설정
            for(int j=0;j<5;j++) {
                if(j == 2)
                    continue;
                xsheet.autoSizeColumn(j);
                xsheet.setColumnWidth(j, (xsheet.getColumnWidth(j))+256);
            }


            // 이미지 넣을 셀 병합 (1번 3번)
            for(int i=0;i<row*4/2;i+=4){
                xsheet.addMergedRegion(new CellRangeAddress(
                        i,
                        i+1,
                        0,
                        0
                ));
                xsheet.addMergedRegion(new CellRangeAddress(
                        i,
                        i+1,
                        3,
                        3
                ));
            }

            // 시트 꾸미기



            xworkbook.write(file);
            file.close();

        } catch(Exception e) {
            e.printStackTrace();
        }
    }
}
