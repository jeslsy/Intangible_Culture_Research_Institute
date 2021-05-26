package com.cics.Reconstructing_Excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;

import java.io.FileInputStream;
import java.io.InputStream;

public class test_image {
    /***
    // 이미지 넣기
    for(int i = 0; i < 30; i++){
        //FileInputStream은 이미지 파일에서 입력 바이트를 얻습니다.
        InputStream inputStream = new FileInputStream("/home/axel/Bilder/"+i+"jpg");

        //InputStream의 내용을 byte []로 가져옵니다.
        byte[] bytes = IOUtils.toByteArray(inputStream);


        //통합 문서에 그림을 추가합니다.
        int pictureIdx = xworkbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);

        //입력 스트림 닫기
        inputStream.close();

        //구체적인 클래스 인스턴스화를 처리하는 객체를 반환합니다.
        CreationHelper helper = xworkbook.getCreationHelper();

        //최상위 도면 족장을 작성합니다.
        Drawing drawing = xsheet.createDrawingPatriarch();

        //워크 시트에 첨부 된 앵커 만들기
        ClientAnchor anchor = helper.createClientAnchor();

        //왼쪽 위 셀 _and_ 오른쪽 아래 셀이있는 앵커 만들기
        anchor.setCol1(1); //Column B
        anchor.setRow1(2); //Row 3
        anchor.setCol2(2); //Column C
        anchor.setRow2(3); //Row 4

        //그림을 만듭니다.
        Picture pict = drawing.createPicture(anchor, pictureIdx);


        // 여기서 if문 넣고 넣어주면 될듯
        cell = xsheet.createRow(2).createCell(1);

    }
     ***/
}
