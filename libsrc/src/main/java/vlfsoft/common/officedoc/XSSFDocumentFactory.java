package vlfsoft.common.officedoc;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class XSSFDocumentFactory {
    private XSSFDocumentFactory() {
    }

    public static XSSFWorkbook getInstance(String aPathname) throws IOException {
        try (FileInputStream in = new FileInputStream(new File(aPathname))) {
            return new XSSFWorkbook(in);
        }
    }

}
