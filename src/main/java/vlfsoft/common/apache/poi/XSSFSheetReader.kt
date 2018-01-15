package vlfsoft.common.apache.poi

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import vlfsoft.concepts.SyntacticSugar
import java.io.File
import java.io.FileInputStream
import java.util.*
import java.util.stream.Stream
import java.util.stream.StreamSupport

/**
 * Extension function
 */
fun <R> File.useWorkbook(block: (workbookFile: File, workbook: XSSFWorkbook) -> R): R =
        FileInputStream(this).use { inputStream ->
            XSSFWorkbook(inputStream).use {
                block(this, it)
            }
        }

fun <R> File.useSheet(index: Int = 0, block: (workbookFile: File, sheet: XSSFSheet) -> R): R =
        useWorkbook { _, workbook ->
            block(this, workbook[index])
        }

class XSSFWorkbookGetSheetException(message: String) : Exception(message)

@SyntacticSugar
operator fun XSSFWorkbook.get(index: Int) = getSheetAt(index) ?: throw XSSFWorkbookGetSheetException("There no sheet for index = $index")
operator fun XSSFWorkbook.get(name: String) = getSheet(name) ?: throw XSSFWorkbookGetSheetException("There no sheet for name = '$name'")

class XSSFSheetGetRowException(message: String) : Exception(message)
fun XSSFSheet.getSafeRow(row: Int) = getRow(row) ?: throw XSSFSheetGetRowException("Row[$row] is null")
operator fun XSSFSheet.get(row: Int) = getSafeRow(row)

fun XSSFSheet.getRowEnd(): Int = lastRowNum.let { if (it > 0) it else physicalNumberOfRows - 1 }
fun XSSFSheet.asSequence(startRowNum: Int = 0, endRowNum: Int = getRowEnd(), endOfStreamPredicate: (Row) -> Boolean = { true }): Sequence<XSSFRow> {
    var rowNum = startRowNum
    return generateSequence {
        if (rowNum <= endRowNum) {
            val r = getSafeRow(rowNum++); if (endOfStreamPredicate(r)) r else null
        } else null
    }
}

// IC doesn't recognize
fun XSSFSheet.asStream(startRowNum: Int = 0, endRowNum: Int = getRowEnd(), endOfStreamPredicate: (Row) -> Boolean = { true }) =
            asSequence(startRowNum, endRowNum, endOfStreamPredicate).asStream()

/**
 * Temporary WO:
 */
fun <T> Sequence<T>.asStream(): Stream<T> = StreamSupport.stream({ Spliterators.spliteratorUnknownSize(iterator(), Spliterator.ORDERED) }, Spliterator.ORDERED, false)

class XSSFSheetGetCellValueException(message: String) : Exception(message)

/**
 * PRB: getCell(col) reads cell with format "общий" as null
 * WO 1: Use getCell(col) instead of getSafeCell (return Cell? instead of throwing the exception in Row.get)
 * WO 2: Use Row.invoke instead of Row.get with stringCellValueWithEmptyAsNull: this[0](1)?.stringCellValueWithEmptyAsNull
 * WO 3: Use fun Row.stringCellValueWithEmptyAsNull (col:Int)
 * instead of Cell.stringCellValueWithEmptyAsNull get() = stringCellValue.run { if (isNotEmpty()) this else null }
 *
 */
fun Row.getSafeCell(col: Int, itemName: String = "") = getCell(col) ?: throw XSSFSheetGetCellValueException("${itemName}value is null at column: $col")

operator fun Row.get(col: Int, itemName: String = "") = getSafeCell(col, itemName)

operator fun Row.invoke(col: Int): Cell? = getCell(col)

val Cell.stringCellValueWithEmptyAsNull get() = stringCellValue.run { if (isNotEmpty()) this else null }
fun Row.stringCellValueWithEmptyAsNull (col:Int) = getCell(col)?.stringCellValue?.run { if (isNotEmpty()) this else null }
