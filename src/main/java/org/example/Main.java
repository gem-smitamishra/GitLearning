package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) throws IOException {
        System.out.println("Hello world!");
       FetchData();
    }


    public static void FetchData() throws IOException {
        FileInputStream file = new FileInputStream(new File("C:/smita/DemoData.xlsx"));
        XSSFWorkbook wb = new XSSFWorkbook(file);
        XSSFSheet sheet = wb.getSheetAt(0);
        Iterator<Row> itr= sheet.iterator();
        while (itr.hasNext()){
            Row row = itr.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()){
                Cell cell = cellIterator.next();
                System.out.println(cell.getCellType());
               System.out.println(cell.getStringCellValue());

            }
        }

    }
}