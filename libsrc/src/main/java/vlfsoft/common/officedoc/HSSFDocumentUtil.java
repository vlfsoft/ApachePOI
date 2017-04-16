package vlfsoft.common.officedoc;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
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

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

public final class HSSFDocumentUtil {
    private HSSFDocumentUtil() {
    }

    public static ColorRGB getColorPattern(HSSFColor color, Boolean aTextColor) {
        short[] triplet = null;

        // http://stackoverflow.com/questions/13260470/how-to-get-excel-cell-font-color-rgb-using-apache-poi
        // I think I might have figured it out. When a cell is using default color (black usually), the default color does not exist in color palette.

        // PRB: getFillBackgroundColor always returns invalid value.
        // WO: Use getFillForegroundColor
        // PRB: If in xls are not applied custom colors then HSSFColor.AUTOMATIC is returned by getFillForegroundColor
        // and null by cellStyle.getFont(xlsBook).getColor()
        // WO: Based on http://stackoverflow.com/questions/39667614/get-cell-colour-with-apache-poi
        if (color == null || color instanceof HSSFColor.AUTOMATIC) {
            triplet = aTextColor ? new HSSFColor.BLACK().getTriplet() : new HSSFColor.WHITE().getTriplet();
        }else {
            triplet = color.getTriplet();
        }

        return new ColorRGB(triplet[0], triplet[1], triplet[2]);
    }

    public static ColorRGB getColorPattern(HSSFPalette aPalette, short aColorIdx, Boolean aTextColor){
        return getColorPattern(aPalette.getColor(aColorIdx), aTextColor);
    }

/*
    public final static class AUTOMATIC extends HSSFColor
    {
        private static HSSFColor instance = new AUTOMATIC();

        public final static short   index     = 0x40;

        public short getIndex()
        {
            return index;
        }

        public short [] getTriplet()
        {
            return BLACK.triplet;
        }

        public String getHexString()
        {
            return BLACK.hexString;
        }

        public static HSSFColor getInstance() {
            return instance;
        }
    }
*/

}