package vlfsoft.common.officedoc;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Optional;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import vlfsoft.common.annotations.design.patterns.CreationalPattern;

/**
 * The service provides methods {@link #getInstance()} to create XWPFDocument using existing docx files.
 */
public class XWPFDocumentFactory {

    private final static String BLANK_FILE = "Blank.docx";

    private final @Nonnull File mFile;

    public XWPFDocumentFactory(final @Nonnull File aFile) {
        mFile = aFile;
    }

    public XWPFDocumentFactory(final @Nonnull String aPathname) {
        this(new File(aPathname));
    }

    public XWPFDocumentFactory() {
        this(BLANK_FILE);
    }

    @CreationalPattern.SimpleFactory
    public @Nonnull XWPFDocument getInstance() throws IOException {
        try (FileInputStream in = new FileInputStream(mFile)) {
            return new XWPFDocument(in);
        }
    }

    public static @Nonnull XWPFDocument getInstance(final @Nonnull File aFile) throws IOException {
        return new XWPFDocumentFactory(aFile).getInstance();
    }

    public static @Nonnull XWPFDocument getInstance(final @Nonnull String aPathname) throws IOException {
        return new XWPFDocumentFactory(aPathname).getInstance();
    }
}
