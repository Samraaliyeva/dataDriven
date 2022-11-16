package org.example;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class dynamicData {
    @Test
    public void testData() throws IOException {
        String path = "C:\\Users\\SamraAliyeva\\Downloads\\EXCEL\\Tariffario_NSTAR_2022_Umbria.xls";
        File newFile=new File(path);
        FileInputStream myFile = new FileInputStream(newFile);
        XSSFWorkbook WB=new XSSFWorkbook(myFile);
        XSSFSheet mySheet= WB.getSheet("Test Case");

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
