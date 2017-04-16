package vlfsoft.common.officedoc;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import vlfsoft.common.annotations.design.patterns.gof.CreationalPattern;

public class XWPFDocumentFactory {

    private final static String BLANK_FILE = "Blank.docx";

    private File mFile;

    public XWPFDocumentFactory() {
        mFile = new File(BLANK_FILE);
    }

    public XWPFDocumentFactory(String aPathname) {
        mFile = new File(aPathname);
    }

    public XWPFDocumentFactory(File aFile) {
        mFile = aFile;
    }

    @CreationalPattern.FactoryMethod
    public XWPFDocument getInstance() throws IOException {
        try (FileInputStream in = new FileInputStream(mFile)) {
            return new XWPFDocument(in);
        }
    }

    public static XWPFDocument getInstance(File aFile) throws IOException {
        return new XWPFDocumentFactory(aFile).getInstance();
    }

    public static XWPFDocument getInstance(String aPathname) throws IOException {
        return new XWPFDocumentFactory(aPathname).getInstance();
    }
}
