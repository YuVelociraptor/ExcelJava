package jp.complexalpha.fastx;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.reader.Cell;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;

import java.io.InputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class XlsxFastx {

    public static void writeNewXlsx(String filePath, int rows) throws IOException {

        try(
                FileOutputStream out = new FileOutputStream(filePath);
                Workbook workbook = new Workbook(out, "app", "1.0");
                ){
            Worksheet ws = workbook.newWorksheet("Sheet001");

            // ヘッダ
            ws.value(0, 0, "Number");

            for (int i = 1; i <= rows; i++) {
                ws.value(i, 0, i);
            }

        }
    }

    public static void readExcel() throws IOException {
        try (
                InputStream in = XlsxFastx.class.getClassLoader().getResourceAsStream("read_test.xlsx")
        ) {
            if (in == null) {
                throw new IOException("Resource not found: read_test.xlsx");
            }

            try (ReadableWorkbook workbook = new ReadableWorkbook(in)) {
                Sheet sheet = workbook.getFirstSheet();

                try (Stream<Row> rows = sheet.openStream()) {

                    rows.forEach(row -> {
                        row.forEach(cell -> {
                            System.out.print(cell.getText() + "\t");
                        });
                        System.out.println();
                    });
                }
            }
        }
    }
}
