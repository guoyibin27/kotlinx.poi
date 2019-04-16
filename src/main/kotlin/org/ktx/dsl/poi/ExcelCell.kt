package org.ktx.dsl.poi

import org.apache.poi.hssf.usermodel.HSSFRichTextString
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook

class ExcelCell : ExcelElement {

    var value: String? = null
    var style: ExcelCellStyle? = null

    fun style(init: ExcelCellStyle.() -> Unit): ExcelCellStyle {
        style = ExcelCellStyle()
        style!!.init()
        return style!!
    }

    fun create(workbook: Workbook, row: Row, cellIndex: Int) {
        val cell = row.createCell(cellIndex)
        cell?.setCellValue(HSSFRichTextString(value))
        val style = workbook.createCellStyle()
        this.style?.getCellStyle(workbook, style)
        cell?.cellStyle = style
    }

}