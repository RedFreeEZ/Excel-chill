package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class DischargeExcel extends Main{
    ArrayList<String> array = new ArrayList<>();


    public void readExcel() throws IOException{

        FileInputStream file = new FileInputStream(new File("C:/HomeWork/Test1.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);//берем данные с листа 1 из эксель
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {// Идти по строке пока не наткнемся на пустое значение
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            List<String> list = new ArrayList<>();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case NUMERIC:
                        if (cell.getCellType() == CellType.NUMERIC) {
                            list.add(NumberToTextConverter.toText(cell.getNumericCellValue()));
                            var a = NumberToTextConverter.toText(cell.getNumericCellValue());
                            boolean result = a.matches("(\\+*)\\d{11}");
                            if (result) {
                                array.add(a );
                            }
                        }break;
                }
            }
            System.out.println();
        }
        file.close();
    }



    public void writeExcell() throws IOException, IOException {
        System.out.println(array);
        ArrayList<String> array1 = array;
        Workbook workbook1 = new XSSFWorkbook();
        Sheet newSheet = workbook1.createSheet("NewSheet");
        for(var i=0;i<array1.size();i++){
            Row row1 = newSheet.createRow(i);
            row1.createCell(0).setCellValue(array1.get(i));
        }
        try {
            FileOutputStream fileOut = new FileOutputStream("C:/HomeWork/Write.xlsx");
            workbook1.write(fileOut);
            fileOut.close();
            System.out.println("Файл создан!!!");
        }
        catch (Exception e) {
            System.out.println("Error");
        }
    }
}


