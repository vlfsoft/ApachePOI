package vlfsoft.common.officedoc;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Optional;

import org.jetbrains.annotations.NotNull;

import vlfsoft.patterns.GOF;

public class XSSFDocumentFactory {

    private Optional<File> mFile;

    @GOF.SimpleFactory
    public XSSFWorkbook getInstance() throws IOException {
        if (mFile.isPresent()) {
            try (FileInputStream in = new FileInputStream(mFile.get())) {
                return new XSSFWorkbook(in);
            }
        }else {
            return new XSSFWorkbook();
        }
    }

    public XSSFDocumentFactory(Optional<File> aFile) {
        this.mFile = aFile;
    }

    public XSSFDocumentFactory() {
        this(Optional.empty());
    }

    public XSSFDocumentFactory(final @NotNull String aPathname) {
        this(Optional.of(new File(aPathname)));
    }

}
