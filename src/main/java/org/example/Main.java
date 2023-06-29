package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;


import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        Document document;
        
        List<ProductData> products = new ArrayList<>();

        for (int page = 1; page < 13; page++) {
            try {
                document = Jsoup.connect("https://www.sulpak.kz/f/noutbuki?page=" + page).get();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }

            Element productsDiv = document.selectFirst("div#products");

            if (productsDiv != null) {
                Elements productElements = productsDiv.select("div.product__item");


                for (Element productElement : productElements) {
                    Element titleElement = productElement.selectFirst("div.product__item-name > a");
                    String title = titleElement != null ? titleElement.text() : "";

                    Element priceElement = productElement.selectFirst("div.product__item-price");
                    String price = (priceElement != null) ? priceElement.ownText().trim() : "";

                    String brand = productElement.attr("data-brand");
                    String code = productElement.attr("data-code");

                    if (!title.isEmpty()) {
                        ProductData productData = new ProductData(title, price, brand, code);
                        products.add(productData);
                    }

                    System.out.println("Item: " + title);
                    System.out.println("Price: " + price);
                    System.out.println("Brand: " + brand);
                    System.out.println("Code: " + code);
                }

            } else {
                System.out.println("Products container div not found in the HTML.");
            }
        }
        writeDataToExcel(products, "output.xlsx");
    }

    private static void writeDataToExcel(List<ProductData> products, String outputFile) {
        try (Workbook workbook = new HSSFWorkbook()) {

            Sheet sheet = workbook.createSheet("Products");

            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Title");
            headerRow.createCell(1).setCellValue("Price");
            headerRow.createCell(2).setCellValue("Brand");
            headerRow.createCell(3).setCellValue("Code");

            int rowNum = 1;
            for (ProductData product : products) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(product.getTitle());
                row.createCell(1).setCellValue(product.getPrice());
                row.createCell(2).setCellValue(product.getBrand());
                row.createCell(3).setCellValue(product.getCode());
            }

            for (int i = 0; i < 4; i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
                workbook.write(outputStream);
            }
            System.out.println("Data written to " + outputFile + " successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
