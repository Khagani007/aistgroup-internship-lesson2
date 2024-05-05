package org.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtility {
    private Workbook workbook;
    private Sheet sheet;

    public void initializeWorkbook() throws IOException {

        File file = new File("CarData.xlsx");
        if (file.exists()) {

            FileInputStream inputStream = new FileInputStream(file);
            workbook = new XSSFWorkbook(inputStream);
            sheet = workbook.getSheetAt(0);
            inputStream.close();
        } else {

            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("Car Data");
        }
    }

    public void extractColumnNames() {

        String[] columnNames = {
                "Şəhər", "Marka", "Model", "Buraxılış ili", "Ban növü",
                "Rəng", "Mühərrik", "Yürüş", "Sürətlər qutusu", "Ötürücü",
                "Yeni", "Yerlərin sayı", "Vəziyyəti", "Hansı bazar üçün yığılıb"
        };


        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < columnNames.length; i++) {
            headerRow.createCell(i).setCellValue(columnNames[i]);
        }
    }

    public void writeWorkbookToFile() throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream("CarData.xlsx")) {
            workbook.write(outputStream);
        } finally {
            workbook.close();
        }
    }

    public Sheet getSheet() {
        return sheet;
    }
}
