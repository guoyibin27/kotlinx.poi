package org.ktx.dsl.poi

import org.apache.poi.hssf.usermodel.HSSFRichTextString
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.usermodel.Workbook

class ExcelTitleRow : ExcelElement {

    var titles: List<String> = listOf()
    var horizontalAlignment: HorizontalAlignment = HorizontalAlignment.CENTER
    var verticalAlignment: VerticalAlignment = VerticalAlignment.CENTER

    override fun create(excelElement: ExcelElement, workbook: Workbook, onSheet: Int?) {
        val sheet = workbook.getSheetAt(onSheet ?: 0)
        val titleRow = sheet.createRow(0)
        titles.forEachIndexed { idx, it ->
            val cell = titleRow?.createCell(idx)
            cell?.setCellValue(HSSFRichTextString(it))
            val style = sheet?.workbook?.createCellStyle()
            style?.alignment = horizontalAlignment
            style?.verticalAlignment = verticalAlignment
            cell?.cellStyle = style
        }
    }
}

