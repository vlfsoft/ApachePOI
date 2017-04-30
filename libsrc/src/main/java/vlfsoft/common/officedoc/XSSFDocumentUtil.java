package vlfsoft.common.officedoc;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class XSSFDocumentUtil {
    private XSSFDocumentUtil() {
    }

    /**
     *
     * @param aXSSfWorkbook -
     * @return - cell style for hyperlinks by default hyperlinks are blue and underlined
     */
    public static CellStyle getHlinkCellStyle(XSSFWorkbook aXSSfWorkbook) {
        CellStyle hlinkCellStyle = aXSSfWorkbook.createCellStyle();
        Font hlink_font = aXSSfWorkbook.createFont();
        hlink_font.setUnderline(Font.U_SINGLE);
        hlink_font.setColor(IndexedColors.BLUE.getIndex());
        hlinkCellStyle.setFont(hlink_font);
        return hlinkCellStyle;
    }

}