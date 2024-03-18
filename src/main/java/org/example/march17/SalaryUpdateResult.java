package org.example.march17;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class SalaryUpdateResult {
    public static void main(String[] args)throws Exception {
        FileInputStream file = new FileInputStream("C:\\Users\\rbsap\\IntelljWorkSpace\\ExamModule\\" +
                "src\\main\\resources\\employee.xlsx");
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);


        for (Row row : sheet) {
            if (row.getRowNum() == 0) {

                continue;
            }

            Cell expCell = row.getCell(4);
            int experience = (int) expCell.getNumericCellValue();

            double salary;
            if (experience < 5) {
                salary = 1000 * 5;
                System.out.println(salary);
            } else if (experience >= 5 && experience < 10) {
                salary = 2500 * 5;
                System.out.println(salary);
            } else if (experience >= 10 && experience < 20) {
                salary = 5000 * 5;
                System.out.println(salary);
            } else {
                salary = 8000 * 5;
                System.out.println(salary);
            }

            Cell salaryCell = row.createCell(row.getLastCellNum());
            salaryCell.setCellValue(salary);
        }


        Row headerRow = sheet.getRow(0);
        Cell newHeaderCell = headerRow.createCell(headerRow.getLastCellNum());
        newHeaderCell.setCellValue("Employee Monthly Salary");


        file.close();
        FileOutputStream outFile = new FileOutputStream("C:\\Users\\rbsap\\IntelljWorkSpace\\ExamModule\\" +
                "src\\main\\resources\\employee.xlsx");
        workbook.write(outFile);
        outFile.close();
        workbook.close();

        System.out.println("Employee monthly salary calculation and update completed successfully.");

    }
}
