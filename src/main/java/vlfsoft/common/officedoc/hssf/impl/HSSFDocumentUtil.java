package vlfsoft.common.officedoc.hssf.impl;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import vlfsoft.common.exceptions.AppException;
import vlfsoft.common.officedoc.ColorRGB;
import vlfsoft.common.util.DateTimeUtilsExtKt;
import vlfsoft.common.util.DatetimeUtils;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import java.time.LocalDate;
import java.util.Locale;

import static vlfsoft.common.util.DatetimeUtils.asLocalDate;

public final class HSSFDocumentUtil {
    private HSSFDocumentUtil() {
    }

    public static int getRowEnd(final @NotNull HSSFSheet aXlsSheet) {
        int rowEnd = aXlsSheet.getLastRowNum();
        if (rowEnd == 0) rowEnd = aXlsSheet.getPhysicalNumberOfRows() - 1;
        return rowEnd;
    }

    public static @Nullable
    String getStringCellValue(final @NotNull Cell aCell) {
        // Ignore depreciation warning
        // Cell with Phone can be NUMERIC, if phone number is without '+'.
        switch (aCell.getCellTypeEnum()) {
            case NUMERIC:
                // https://stackoverflow.com/questions/3148535/how-to-read-excel-cell-having-date-with-apache-poi
                if (HSSFDateUtil.isCellDateFormatted(aCell)) {
                    // 42942 - 26.07.2017
                    return DateTimeUtilsExtKt.toString(asLocalDate(aCell.getDateCellValue()), DateTimeUtilsExtKt.ANSI_DATE_FORMAT);
                }else {
                    return String.valueOf((int) aCell.getNumericCellValue());
                }
            case STRING:
                return aCell.getStringCellValue();

        }
        return null;
    }

    /**
     * https://stackoverflow.com/questions/5794659/poi-how-do-i-set-cell-value-to-date-and-apply-default-excel-date-format
     */
    public static void setDateValue(final @NotNull HSSFWorkbook aWorkbook,  final @NotNull Cell aCell, final @NotNull LocalDate aDate, final @NotNull String aFormat) {
        CellStyle cellStyle = aWorkbook.createCellStyle();
        CreationHelper createHelper = aWorkbook.getCreationHelper();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(aFormat));
        aCell.setCellValue(DatetimeUtils.asDate(aDate));
        aCell.setCellStyle(cellStyle);
    }

    private static ColorRGB getColorPattern(HSSFColor color, Boolean aTextColor) {
        short[] triplet;

        // http://stackoverflow.com/questions/13260470/how-to-get-excel-cell-font-color-rgb-using-apache-poi
        // I think I might have figured it out. When a cell is using default color (black usually), the default color does not exist in color palette.

        // PRB: getFillBackgroundColor always returns invalid value.
        // WO: Use getFillForegroundColor
        // PRB: If in xls are not applied custom colors then HSSFColor.AUTOMATIC is returned by getFillForegroundColor
        // and null by cellStyle.getFont(xlsBook).getColor()
        // WO: Based on http://stackoverflow.com/questions/39667614/get-cell-colour-with-apache-poi
        if (color == null || color instanceof HSSFColor.AUTOMATIC) {
            triplet = aTextColor ? new HSSFColor.BLACK().getTriplet() : new HSSFColor.WHITE().getTriplet();
        } else {
            triplet = color.getTriplet();
        }

        return new ColorRGB(triplet[0], triplet[1], triplet[2]);
    }

    public static ColorRGB getColorPattern(HSSFPalette aPalette, short aColorIdx, Boolean aTextColor) {
        return getColorPattern(aPalette.getColor(aColorIdx), aTextColor);
    }

    static
    public
    @NotNull
    String getDataItem(final @NotNull Row aXlsRow, int aCol,
                       final @NotNull String aDataItemName,
                       boolean aCheckEmpty,
                       final @NotNull AppException.ErrorCodeAdapterA aErrorCode) throws AppException {
        Cell cell = aXlsRow.getCell(aCol);

        if (cell == null) {
            if (aCheckEmpty) {
                AppException.newExceptionWithBriefMessage(aErrorCode,
                        String.format(Locale.getDefault(), "%s is null at aCol: %d", aDataItemName, aCol), "");
            }
            return "";
        }

        @Nullable String dataItem = HSSFDocumentUtil.getStringCellValue(cell);

        if (dataItem == null) {
            if (aCheckEmpty) {
                AppException.newExceptionWithBriefMessage(aErrorCode,
                        String.format(Locale.getDefault(), "%s is null at aCol: %d", aDataItemName, aCol), "");
            }
            return "";
        }

        if (aCheckEmpty && dataItem.isEmpty()) {
            AppException.newExceptionWithBriefMessage(aErrorCode,
                    String.format(Locale.getDefault(), "%s is empty at aCol: %d", aDataItemName, aCol), "");
            return "";
        }

        return dataItem;
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