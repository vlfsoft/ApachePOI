package vlfsoft.common.officedoc;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

final public class XWPFDocumentWriter {
    private XWPFDocumentWriter() {
    }

    public static void toDocx(XWPFDocument aDocument, String aPathname) throws IOException {
        toDocx(aDocument, new File(aPathname));
    }

    public static void toDocx(XWPFDocument aDocument, File aFile) throws IOException {
        try (OutputStream out = new FileOutputStream(aFile)) {
            aDocument.write(out);
        }
    }

    public static void toPdf(XWPFDocument aDocument, File aFile) throws IOException {
        try (OutputStream out = new FileOutputStream(aFile)) {
            PdfConverter.getInstance().convert(aDocument, out, PdfOptions.create());
        }
    }

    public static void toPdf(XWPFDocument aDocument, String aPathname) throws IOException {
        toPdf(aDocument, new File(aPathname));
    }

    public static void toPdf(File aFileDocx, File aFile) throws IOException {
        toPdf(XWPFDocumentFactory.getInstance(aFileDocx), aFile);
    }

    public static void toPdf(String aPathnameDocx, String aPathname) throws IOException {
        toPdf(new File(aPathnameDocx), new File(aPathname));
    }

}
