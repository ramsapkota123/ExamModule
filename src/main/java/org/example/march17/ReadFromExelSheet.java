package org.example.march17;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class ReadFromExelSheet {
    public static void main(String[] args) throws Exception{
        File file = new File("C:\\Users\\rbsap\\IntelljWorkSpace\\ExamModule\\" +
                "src\\main\\resources\\employee.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        System.out.println("binary format data: "+fileInputStream);

        XSSFWorkbook xssfWorkbook=new XSSFWorkbook(fileInputStream);
        int numberOfSheets=xssfWorkbook.getNumberOfSheets();
        System.out.println("number of sheets"+numberOfSheets);

        for(int i=0;i<numberOfSheets;i++){
            XSSFSheet xssfSheet=xssfWorkbook.getSheetAt(i);
            System.out.println(xssfSheet.getPhysicalNumberOfRows());

            Iterator<Row> rowIterator =xssfSheet.iterator();
            while (rowIterator.hasNext()){
                Row row=rowIterator.next();
                System.out.println(row.getPhysicalNumberOfCells());

                Iterator<Cell> cellIterator =row.cellIterator();
                while (cellIterator.hasNext()){
                    Cell cell=cellIterator.next();
                    switch (cell.getCellType()){
                        case STRING :
                            System.out.println(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            System.out.println(cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            System.out.println(cell.getBooleanCellValue());
                            break;

                    }

                }
            }


        }



    }
}



