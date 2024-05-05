package org.example;

import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        int startPage = 1;
        int endPage = 5;


        ExcelUtility excelUtility = new ExcelUtility();
        excelUtility.initializeWorkbook();
        excelUtility.extractColumnNames();


        ScrapingUtility scrapingUtility = new ScrapingUtility(excelUtility.getSheet());
        for (int page = startPage; page <= endPage; page++) {
            scrapingUtility.scrapeAndWriteData(page);
        }

        excelUtility.writeWorkbookToFile();
    }
}
