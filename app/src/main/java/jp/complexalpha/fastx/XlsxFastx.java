package jp.complexalpha.fastx;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class XlsxFastx {

    public static void writeXlsx(String filePath, int rows) throws IOException {

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
}
