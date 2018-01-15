package vlfsoft.common.officedoc;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.jetbrains.annotations.NotNull;

import vlfsoft.patterns.GOF;

/**
 * The service provides methods {@link #getInstance()} to create XWPFDocument using existing docx files.
 */
public class XWPFDocumentFactory {

    private final static String BLANK_FILE = "Blank.docx";

    private final @NotNull File mFile;

    public XWPFDocumentFactory(final @NotNull File aFile) {
        mFile = aFile;
    }

    public XWPFDocumentFactory(final @NotNull String aPathname) {
        this(new File(aPathname));
    }

    public XWPFDocumentFactory() {
        this(BLANK_FILE);
    }

    @GOF.Factory.SimpleFactory
    public @NotNull XWPFDocument getInstance() throws IOException {
        try (FileInputStream in = new FileInputStream(mFile)) {
            return new XWPFDocument(in);
        }
    }

    public static @NotNull XWPFDocument getInstance(final @NotNull File aFile) throws IOException {
        return new XWPFDocumentFactory(aFile).getInstance();
    }

    public static @NotNull XWPFDocument getInstance(final @NotNull String aPathname) throws IOException {
        return new XWPFDocumentFactory(aPathname).getInstance();
    }
}
