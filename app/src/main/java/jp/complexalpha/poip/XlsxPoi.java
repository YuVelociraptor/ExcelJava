package jp.complexalpha.poip;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class XlsxPoi {

    public void writeNewXlsx(String filePath) throws IOException {


        try (
                Workbook workbook = new XSSFWorkbook();
                FileOutputStream out = new FileOutputStream(filePath);
        ) {

            Sheet sheet = workbook.createSheet("シート001");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Number");

            for (int i = 1; i <= 10; i++) {
                Row row = sheet.createRow(i);
                row.createCell(0).setCellValue(i);
            }

            workbook.write(out);
        }
    }
}
