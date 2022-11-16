package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class add {


    @Test
    public void testCase() throws IOException {


        FileInputStream fis = new FileInputStream("C:\\Users\\SamraAliyeva\\Downloads\\EXCEL\\Test Tariffario Umbria.xlsm");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        // Iterate each row one by one
        Iterator<Row> rIterator = sheet.iterator();
        while (rIterator.hasNext())
        {
            Row row = rIterator.next();

            // For each row, iterate through all the columns
            Iterator<Cell> Cell = row.cellIterator();

            while (Cell.hasNext())
            {
                Cell cell = Cell.next();

                // Check the cell type
                switch(cell.getCellType())
                {
                    case STRING:
                        System.out.print(cell.getStringCellValue());
                        break;

                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue());
                        break;

                    case FORMULA:
                        System.out.print(cell.getNumericCellValue());
                        break;
                }
                System.out.print("|");
            }
            System.out.println();
        }
        workbook.close();
        fis.close();
    }
}