package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.IOException;

public class ScrapingUtility {
    private Sheet sheet;

    public ScrapingUtility(Sheet sheet) {
        this.sheet = sheet;
    }

    public void scrapeAndWriteData(int page) {
        String url = "https://turbo.az/?page=" + page;

        try {
            Document document = Jsoup.connect(url).get();
            Elements carLinks = document.select(".products-i__link");

            for (Element carLink : carLinks) {
                String carUrl = carLink.absUrl("href");
                extractCarData(carUrl);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void extractCarData(String url) {
        try {
            Document document = Jsoup.connect(url).get();
            Elements productProperties = document.select(".page-content");

            if (productProperties.isEmpty()) {
                System.out.println("No product properties found for URL: " + url);
                return;
            }

            Row row = sheet.createRow(sheet.getLastRowNum() + 1);

            for (Element propertyItem : productProperties) {
                Elements propertyNames = propertyItem.select(".product-properties__i-name");
                Elements propertyValues = propertyItem.select(".product-properties__i-value");

                int cellNum = 0;

                for (int i = 0; i < propertyNames.size(); i++) {
                    String propertyName = propertyNames.get(i).text();
                    String propertyValue = propertyValues.get(i).text();
                    cellNum = findOrCreateCell(propertyName, cellNum);

                    row.createCell(cellNum).setCellValue(propertyValue);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private int findOrCreateCell(String propertyName, int startCellNum) {
        Row headerRow = sheet.getRow(0);


        if (headerRow == null) {
            headerRow = sheet.createRow(0);
        }

        int lastCellNum = headerRow.getLastCellNum();
        int cellNum;


        for (cellNum = 0; cellNum < lastCellNum; cellNum++) {
            Cell cell = headerRow.getCell(cellNum);
            if (cell != null && cell.getStringCellValue().equals(propertyName)) {
                return cellNum;
            }
        }


        headerRow.createCell(lastCellNum).setCellValue(propertyName);
        return lastCellNum;
    }
}
