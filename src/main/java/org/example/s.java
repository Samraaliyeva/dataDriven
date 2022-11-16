package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class s {

    @Test
    public void testCase() throws IOException {

       // File file = new File("C:\\Users\\SamraAliyeva\\Downloads\\EXCEL\\Test Tariffario Umbria.xlsm");
        //FileInputStream inputStream = new FileInputStream(file);
        //HSSFWorkbook wb=new HSSFWorkbook(inputStream);
        FileInputStream fis = new FileInputStream("C:\\Users\\SamraAliyeva\\Downloads\\EXCEL\\Test Tariffario Umbria.xlsm");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //I have added test data in the cell A1 as "SoftwareTestingMaterial.com"
        //Cell A1 = row 0 and column 0. It reads first row as 0 and Column A as 0.
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        System.out.println(cell);
        System.out.println(sheet.getRow(0).getCell(0));

    }
}

