package vlfsoft.common.officedoc;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

final public class XSSFDocumentWriter {
    private XSSFDocumentWriter() {
    }

    public static void toXlsx(XSSFWorkbook aDocument, String aPathname) throws IOException {
        toXlsx(aDocument, new File(aPathname));
    }

    public static void toXlsx(XSSFWorkbook aDocument, File aFile) throws IOException {
        try (OutputStream out = new FileOutputStream(aFile)) {
            aDocument.write(out);
        }
    }

}
