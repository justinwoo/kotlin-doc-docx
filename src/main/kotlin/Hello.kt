package hello

import java.io.FileOutputStream
import java.io.IOException
import java.util.Arrays
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook

@Throws(IOException::class)
fun main(args: Array<String>?) {
    val workbook = XSSFWorkbook()
    val createHelper = workbook.getCreationHelper()
    val sheet = workbook.createSheet("Hello")

    for (y in 0..10) {
        val row = sheet.createRow(y)
        for (x in 0..10) {
            row.createCell(x).setCellValue(x.toString())
        }
    }

    val fileOut = FileOutputStream("hello.xlsx")
    workbook.write(fileOut)
    fileOut.close()
    workbook.close()
}
