package vlfsoft.common.apache.poi

import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Row
import java.io.File
import java.io.FileInputStream
import kotlin.streams.asStream

/**
 * Extension function
 */

fun <R> File.useWorkbookH(block: (workbook: HSSFWorkbook) -> R): R =
        FileInputStream(this).use { inputStream ->
            HSSFWorkbook(inputStream).use {
                block(it)
            }
        }

fun <R> File.useSheet(index: Int = 0, block: (sheet: HSSFSheet) -> R): R =
        useWorkbookH {
            block(it.getSheetAt(index))
        }

class HSSFSheetGetRowException(message: String) : Exception(message)
fun HSSFSheet.getSafeRow(row: Int) = getRow(row) ?: throw HSSFSheetGetRowException("Row[$row] is null")

fun HSSFSheet.getRowEnd(): Int = lastRowNum.let { if (it > 0) it else physicalNumberOfRows - 1 }
fun HSSFSheet.asSequence(aStartRowNum: Int = 0, aEndRowNum: Int = getRowEnd(), aEndOfStreamFunction: (Row) -> Boolean = { true }): Sequence<Row> {
    var rowNum = aStartRowNum
    return generateSequence {
        if (rowNum <= aEndRowNum) {
            val r = getSafeRow(rowNum++); if (aEndOfStreamFunction(r)) r else null
        } else null
    }
}

fun HSSFSheet.asStream(aStartRowNum: Int = 0, aEndRowNum: Int = getRowEnd(), aEndOfStreamFunction: (Row) -> Boolean = { true }) =
        asSequence(aStartRowNum, aEndRowNum, aEndOfStreamFunction).asStream()

//class HSSFSheetGetCellValueException(message: String) : Exception(message)
// fun Row.getSafeCell(col: Int, itemName: String = "") = getCell(col) ?: throw HSSFSheetGetCellValueException("${itemName}value is null at column: $col")


