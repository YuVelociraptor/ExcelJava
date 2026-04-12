package jp.complexalpha.fastx;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class XlsxFastx {

    public void writeXlsx(String filePath) throws IOException {

        try(
                FileOutputStream out = new FileOutputStream(filePath);
                Workbook workbook = new Workbook(out, "", "");
                ){
            Worksheet ws = workbook.newWorksheet("Sheet1");

            // ヘッダ
            ws.value(0, 0, "名前");
            ws.value(0, 1, "年齢");

            // データ
            ws.value(1, 0, "田中");
            ws.value(1, 1, 20);

            ws.value(2, 0, "佐藤");
            ws.value(2, 1, 22);
        }
    }
}
