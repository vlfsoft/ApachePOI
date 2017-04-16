package vlfsoft.common.officedoc;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class XWPFDocumentWriter {
    private XWPFDocumentWriter() {
    }

    public static void toDocx(XWPFDocument aDocument, String aPathnameDocx) throws IOException {
        toDocx(aDocument, aPathnameDocx);
    }

    public static void toDocx(XWPFDocument aDocument, File aFileDocx) throws IOException {
        try (OutputStream out = new FileOutputStream(aFileDocx)) {
            aDocument.write(out);
        }
    }

    public static void toPdf(XWPFDocument aDocument, File aFilePdf) throws IOException {
        try (OutputStream out = new FileOutputStream(aFilePdf)) {
            PdfConverter.getInstance().convert(aDocument, out, PdfOptions.create());
        }
    }

    public static void toPdf(XWPFDocument aDocument, String aPathnamePdf) throws IOException {
        toPdf(aDocument, new File(aPathnamePdf));
    }

    public static void toPdf(File aFileDocx, File aFilePdf) throws IOException {
        toPdf(XWPFDocumentFactory.getInstance(aFileDocx), aFilePdf);
    }

    public static void toPdf(String aPathnameDocx, String aPathnamePdf) throws IOException {
        toPdf(new File(aPathnameDocx), new File(aPathnamePdf));
    }

}
