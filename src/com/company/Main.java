package com.company;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Iterator;

public class Main {

    private static Sheet sheet;

    public static void main(String[] args) {
	// write your code here
        downloadFile();
    }
    public static void downloadFile(){
        URL website = null;

        try {
            website = new URL("http://rksi.ru/rasp/2017.xls");
        } catch (MalformedURLException e) {
            e.printStackTrace();
        }


        try (InputStream in = website.openStream()) {
            Path target = Paths.get("D:\\Raspisanie\\2017.xlsx");
            Files.copy(in, target, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            e.printStackTrace();
        }
        xlsxReader();
    }
    public static void xlsxReader(){
        String excelFilePath = "D:\\Raspisanie\\2017.xlsx";
        FileInputStream inputStream = null;
        try {
            inputStream = new FileInputStream(new File(excelFilePath));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        Workbook workbook = null;
        try {
            workbook = new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        sheet = workbook.getSheetAt(2);
        Iterator<Row> iterator = sheet.iterator();
        int rows=0;
        while(iterator.hasNext()){

            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if(cell.getColumnIndex()==0 || cell.getColumnIndex()==1 || cell.getColumnIndex()==44 || cell.getColumnIndex()==46){
                    Row row= sheet.getRow(rows);
                    if(cell.getColumnIndex()==0 && row.getCell(0)!=null &&!row.getCell(0).getStringCellValue().equals("")){
                        System.out.println(cell.getStringCellValue());

                    }
                    if(row.getCell(44)!=null &&!row.getCell(44).getStringCellValue().equals("") && cell.getColumnIndex()!=0) {
                        System.out.print("   " + cell.getStringCellValue() + "   ");
                    }

                }
            }
            System.out.println("");
            rows++;
        }
    }
}
