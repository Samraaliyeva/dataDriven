package org.example;

import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;

public class read {

    @Test
public void testData() throws IOException {

        String path="C:\\Users\\SamraAliyeva\\Downloads\\EXCEL\\Test Tariffario Umbria.xlsm";
        Workbook workbook= WorkbookFactory.create(new File(path));
        Sheet mySheet=workbook.getSheet("Test Case");
        int rowCount=mySheet.getLastRowNum()-mySheet.getFirstRowNum();

        for (int i=0; i<=rowCount; i++){
            int cellCount=mySheet.getRow(i).getLastCellNum();

            System.out.println("row data: "+i);
            for (int j=0; j<=cellCount; j++){
                System.out.println(mySheet.getRow(i).getCell(j).getStringCellValue() +" ,");
            }
            System.out.println();
        }
    }


    }

