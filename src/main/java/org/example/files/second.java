package org.example.files;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

public class second {

    public static void main(String[] args) throws IOException {


        HSSFWorkbook workbookb = new HSSFWorkbook();
        HSSFSheet sh = workbookb.createSheet("Test Sheet Number 1");

        sh.createRow(0);
        sh.getRow(0).createCell(0).setCellValue("Hello");
        sh.getRow(0).createCell(1).setCellValue("World");

        sh.createRow(1);
        sh.getRow(1).createCell(0).setCellValue("Ciao");
        sh.getRow(1).createCell(1).setCellValue("Mondo");

        File f=new File("C:\\Users\\SamraAliyeva\\Downloads\\EXCEL\\Test1.xls");

        workbookb.write(f);
        workbookb.close();

    }
}