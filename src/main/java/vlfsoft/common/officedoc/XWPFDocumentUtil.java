package vlfsoft.common.officedoc;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;

import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

public final class XWPFDocumentUtil {
    private XWPFDocumentUtil() {
    }

    public static void createTOC(XWPFDocument aDocumnet) {
        // http://stackoverflow.com/questions/40235909/creating-a-table-of-contents-for-a-xwpfdocument-with-page-numbers-indication/40264237#40264237
        XWPFParagraph paragraph = aDocumnet.createParagraph();
        CTP ctP = paragraph.getCTP();
        CTSimpleField toc = ctP.addNewFldSimple();
        // http://www.techrepublic.com/article/use-words-toc-field-to-fine-tune-your-table-of-contents/
        // https://blogs.technet.microsoft.com/tasush/2011/06/15/word-2010-2/
        // \h is the hyperlink switch that turns each entry into a hyperlink to the associated section. If you delete this switch, only the page numbers are hyperlinked.
        // Works only with MS Word
        toc.setInstr("TOC \\h");
        toc.setDirty(STOnOff.TRUE);
    }

    public static void setText1251(XWPFRun aRun, String aText) {
        aRun.setText(vlfsoft.common.util.StringUtils.getString1251(aText));
    }

    public static void appendExternalHyperlink(String url, String text, XWPFDocument aDocumnet) {
        appendExternalHyperlink(url, text, aDocumnet.createParagraph());
    }

    /**
     * http://stackoverflow.com/questions/37928363/how-to-add-a-hyperlink-to-a-xwpfrun
     * @param url -
     * @param text -
     * @param paragraph -
     */
    public static void appendExternalHyperlink(String url, String text, XWPFParagraph paragraph) {

        // Add the link as External relationship
        String id = paragraph.getDocument().getPackagePart().addExternalRelationship(url, XWPFRelation.HYPERLINK.getRelation()).getId();

        //Append the link and bind it to the relationship
        CTHyperlink cLink = paragraph.getCTP().addNewHyperlink();
        cLink.setId(id);

        //Create the linked text
        CTText ctText = CTText.Factory.newInstance();
        ctText.setStringValue(text);
        CTR ctr = CTR.Factory.newInstance();
        ctr.setTArray(new CTText[]{ctText});
        CTRPr rpr = ctr.addNewRPr();
        CTColor colour = CTColor.Factory.newInstance();
        colour.setVal("0000FF");
        rpr.setColor(colour);
        CTRPr rpr1 = ctr.addNewRPr();
        rpr1.addNewU().setVal(STUnderline.SINGLE);

        //Insert the linked text into the link
        cLink.setRArray(new CTR[]{ctr});
    }

    public static void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);

        // is a null op if already defined
        XWPFStyles styles = docxDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);

    }

    public static void addPageBreak(@NotNull final XWPFDocument aDocument) {
        XWPFParagraph paragraph = aDocument.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.addBreak(BreakType.PAGE);
    }

    public static XWPFParagraph addTextWithStyleId(@NotNull final XWPFDocument aDocument,
                                                   @NotNull final String aStyleId,
                                                   @NotNull final String aText,
                                                   boolean aAddCarriageReturn) {
        XWPFParagraph paragraph = aDocument.createParagraph();

        XWPFRun run = paragraph.createRun();
        run.setText(aText);

        if (aStyleId != null) paragraph.setStyle(aStyleId);

        if (aAddCarriageReturn) run.addCarriageReturn();

        return paragraph;
    }

    public static XWPFParagraph addTextWithStyleId(@NotNull final XWPFDocument aDocument,
                                                   @NotNull final String aStyleId,
                                                   @NotNull final String aText) {
        return addTextWithStyleId(aDocument, aStyleId, aText, false);
    }

    public static XWPFParagraph addText(@NotNull final XWPFDocument aDocument,
                                        @NotNull final String aText,
                                        boolean aAddCarriageReturn) {
        return addTextWithStyleId(aDocument, null, aText, aAddCarriageReturn);
    }

    public static XWPFParagraph addText(@NotNull final XWPFDocument aDocument,
                                        @NotNull final String aText) {
        return addText(aDocument, aText, true);
    }

    public static String transferStyle(XWPFStyles aBlankDocumentStyles, XWPFDocument aStylesDocument,
                                       XWPFStyles aStylesDocumentStyles, int aPos) {
        XWPFParagraph paragraph = aStylesDocument.getParagraphArray(aPos);
        String styleId = paragraph.getStyle();

        XWPFStyle style = aStylesDocumentStyles.getStyle(styleId);
        aBlankDocumentStyles.addStyle(style);
        return styleId;
    }

    public static XWPFRun addPicture(XWPFRun aRun, int aPictureType, String aPictureFilename, int aWidth, int aHeight) throws InvalidFormatException, IOException {
        // http://stackoverflow.com/questions/26764889/how-to-insert-a-image-in-word-document-with-apache-poi

        // 18.04.16 Changed from aRun.addPicture(new FileInputStream(aPictureFilename), aPictureType, aPictureFilename, Units.toEMU(aWidth), Units.toEMU(aHeight));
        // to
        aRun.addPicture(new FileInputStream(aPictureFilename), aPictureType, aPictureFilename, Units.pixelToEMU(aWidth), Units.pixelToEMU(aHeight));

        return aRun;
    }

    public static XWPFRun addPicture(XWPFDocument aDocument, int aPictureType, String aPictureFilename, int aWidth, int aHeight) throws InvalidFormatException, IOException {
        XWPFParagraph paragraph = aDocument.createParagraph();
        XWPFRun run = paragraph.createRun();

        return addPicture(run, aPictureType, aPictureFilename, aWidth, aHeight);
    }

    /**
     * See <a href="http://stackoverflow.com/questions/35419619/how-can-i-set-background-colour-of-a-run-a-word-in-line-or-a-paragraph-in-a-do">How can I set background colour of a runAndCatch (a word in line or a paragraph) in a docx filemodified by using Apache POI?</a>
     */
    public static void setBackgroundColor(XWPFRun aRun, String aBackgroundColor) {
        CTShd cTShd = aRun.getCTR().addNewRPr().addNewShd();
        cTShd.setVal(STShd.CLEAR);
        cTShd.setColor("auto");
        cTShd.setFill(aBackgroundColor);
    }

    /**
     * @param aRun            -
     * @param aHighlightColor - It's possible to use STHighlightColor.Enum.forString("yellow"). see {@link STHighlightColor}
     *                        "black" "blue" "cyan" "green" "magenta" "red" "yellow" "white" "darkBlue" "darkCyan" "darkGreen" "darkMagenta" "darkRed" "darkYellow" "darkGray" "lightGray" "none"
     */
    public static void setHighlightColor(XWPFRun aRun, STHighlightColor.Enum aHighlightColor) {
        aRun.getCTR().addNewRPr().addNewHighlight().setVal(aHighlightColor);
    }

    /**
     * See <a href="http://stackoverflow.com/questions/40345285/removing-an-xwpfparagraph-keeps-the-paragraph-symbol-for-it">http://stackoverflow.com/questions/40345285/removing-an-xwpfparagraph-keeps-the-paragraph-symbol-for-it</a>
     * @param aDocument -
     */
    public static void removeAllParagraphs(@NotNull final XWPFDocument aDocument) {
        for (int pNumber = aDocument.getParagraphs().size() - 1; pNumber >= 0; pNumber--) {
            XWPFParagraph p = aDocument.getParagraphs().get(pNumber);
            int pPos = aDocument.getPosOfParagraph(p);
            aDocument.getDocument().getBody().removeP(pPos);
        }

    }
}