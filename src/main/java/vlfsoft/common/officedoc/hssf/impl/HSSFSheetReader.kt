package vlfsoft.common.officedoc.hssf.impl

import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.slf4j.Logger
import vlfsoft.common.annotations.design.patterns.CreationalPattern
import vlfsoft.common.exceptions.AppException
import java.io.File
import java.io.FileInputStream
import java.util.stream.Stream
import kotlin.streams.asStream

class HSSFSheetReader {

    @CreationalPattern.Singleton
    companion object {
        fun useWorkbook(aFile: File, aReaderFunction: (aHSSFWorkbook: HSSFWorkbook) -> Unit) {
            FileInputStream(aFile).use { aInputStream ->
                HSSFWorkbook(aInputStream).use {
                    aReaderFunction(it)
                }
            }
        }

        fun useSheet(aFile: File, aIndex: Int, aReaderFunction: (aHSSFSheet: HSSFSheet) -> Unit) {
            useWorkbook(aFile) {
                aReaderFunction(it.getSheetAt(aIndex))
            }
        }

    }
}

/**
 * Extension function
 * @return null if rowNum < aEndRowNum > aEndRowNum or getRow(rowNum++) returns null
 */
fun HSSFSheet.asStream(aStartRowNum: Int = 0, aEndRowNum: Int = getRowEnd(), aEndOfStreamFunction: (Row) -> Boolean = { true }): Stream<Row> {
    var rowNum = aStartRowNum
    return generateSequence { if (rowNum <= aEndRowNum) {val r = getRow(rowNum++); if (aEndOfStreamFunction(r)) r else null} else null }.asStream()
}

/**
 * Extension functions
 */
fun HSSFSheet.getRowEnd(): Int = HSSFDocumentUtil.getRowEnd(this);
fun Row.cellValueIsNotEmpty(aCol: Int) = getCell(aCol)?.getStringValue()?.isNotEmpty() ?: false
fun Cell.getStringValue() = HSSFDocumentUtil.getStringCellValue(this)
fun Row.getDataItem(aCol: Int, aDataItemName: String, aCheckEmpty: Boolean, log: Logger, aErrorCode: AppException.ErrorCodeAdapterA
): String {
    try {
        return HSSFDocumentUtil.getDataItem(this, aCol, aDataItemName, aCheckEmpty, aErrorCode)
    } catch (e: Exception) {
        log.info("Exception at aCol: {}", aCol)
        AppException.propagate(aErrorCode, e)
    }

    return ""
}
