package com.vermeg.abenabbes;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSXToXLSConverter {

    public static void main(String[] args) throws IOException {
        Scanner sc = new Scanner(System.in);
        System.out.println("This is a converter from XLSX to XLS: ");
        System.out.println("file name + extension(xlsx): \n");
        String inputFile = sc.next();
        String outputFile = "output.xls";

        // Read the XLSX file into a XSSFWorkbook object
        FileInputStream xlsxFile = new FileInputStream(inputFile);
        XSSFWorkbook xlsxWorkbook = new XSSFWorkbook(xlsxFile);

        // Create a new HSSFWorkbook object to write the XLS data to
        HSSFWorkbook xlsWorkbook = new HSSFWorkbook();

        // Delete the default sheet in the new workbook
        Sheet sheet = xlsWorkbook.getSheet("Sheet1");
        if (sheet != null) {
            xlsWorkbook.removeSheetAt(xlsWorkbook.getSheetIndex(sheet));
        }

        int totalSheets = xlsxWorkbook.getNumberOfSheets();
        int progress = 0;
        System.out.print("Conversion progress: ");

        // Loop through each sheet in the XLSX file and copy it to the new workbook
        for (int i = 0; i < totalSheets; i++) {
            XSSFSheet xlsxSheet = xlsxWorkbook.getSheetAt(i);
            Sheet xlsSheet = xlsWorkbook.createSheet(xlsxSheet.getSheetName());
            copySheet(xlsxSheet, xlsSheet);
            progress++;
            int percent = progress * 100 / totalSheets;
            int numberOfHashes = progress * 20 / totalSheets;
            System.out.print(String.format("\r[%-20s] %d%%", new String(new char[numberOfHashes]).replace('\0', '#'), percent));
        }
        System.out.println();

        // Write the XLS data to a file
        FileOutputStream xlsFile = new FileOutputStream(outputFile);
        xlsWorkbook.write(xlsFile);
        xlsFile.close();

        // Close the input and output streams
        xlsxFile.close();
        System.out.println("Conversion done successfully (check output.xls)");
    }

    private static void copySheet(Sheet sourceSheet, Sheet targetSheet) {
        int index = 0;
        for (Row sourceRow : sourceSheet) {
            Row targetRow = targetSheet.createRow(index++);
            for (Cell sourceCell : sourceRow) {
                Cell targetCell = targetRow.createCell(sourceCell.getColumnIndex(), sourceCell.getCellType());
                switch (sourceCell.getCellType()) {
                    case STRING:
                        targetCell.setCellValue(sourceCell.getRichStringCellValue().getString());
                        break;
                    case NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        targetCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    default:
                        targetCell.setCellValue("");
                }
            }
        }
    }
}
