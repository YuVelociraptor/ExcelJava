package jp.complexalpha;

import jp.complexalpha.fastx.XlsxFastx;
import jp.complexalpha.poip.XlsxPoi;

import java.io.IOException;

public class ReadApp {
    public static void main(String[] args) throws IOException {
        System.out.println("ReadApp START");
        //XlsxFastx.readExcel();
        System.out.println("ReadApp END");

        XlsxPoi.readXlsx();
    }
}
