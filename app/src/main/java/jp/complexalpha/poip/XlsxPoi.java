package jp.complexalpha.poip;

import jp.complexalpha.fastx.XlsxFastx;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

import static org.apache.poi.ss.usermodel.CellType.BLANK;
import static org.apache.poi.ss.usermodel.ConditionType.FORMULA;
import static org.apache.poi.ss.usermodel.CellType.STRING;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.BOOLEAN;

public class XlsxPoi {

    public static void writeNewXlsx(String filePath, int rows) throws IOException {

        try (
                Workbook workbook = new XSSFWorkbook();
                FileOutputStream out = new FileOutputStream(filePath);
        ) {

            Sheet sheet = workbook.createSheet("シート001");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Number");

            for (int i = 1; i <= rows; i++) {
                Row row = sheet.createRow(i);
                row.createCell(0).setCellValue(i);
            }

            workbook.write(out);
        }
    }

    public static void readXlsx() throws IOException {
        try (
                InputStream in = XlsxFastx.class.getClassLoader().getResourceAsStream("read_test.xlsx");
        ) {
            if (in == null) {
                throw new IOException("Resource not found: read_test.xlsx");
            }

            try(Workbook workbook = new XSSFWorkbook(in)){
                Sheet sheet = workbook.getSheetAt(0);

                for(Row row : sheet) {
                    for(Cell cell : row) {
                        switch (cell.getCellType()) {
                            case STRING:
                                System.out.print(cell.getStringCellValue() + "\t");
                                break;

                            case NUMERIC:
                                System.out.print(cell.getNumericCellValue() + "\t");
                                break;

                            case BOOLEAN:
                                System.out.print(cell.getBooleanCellValue() + "\t");
                                break;

                            case FORMULA:
                                System.out.print(cell.getCellFormula() + "\t");
                                break;

                            case BLANK:
                                System.out.print("\t");
                                break;

                            default:
                                System.out.print("\t");
                        }
                    }

                    System.out.println();
                }
            }

        }
    }
}
