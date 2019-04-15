package org.ktx.dsl.poi

import org.apache.poi.hssf.usermodel.HSSFRichTextString
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.usermodel.Workbook

class ExcelRow : ExcelElement {

    var data: List<List<Any>>? = null
    var horizontalAlignment: HorizontalAlignment = HorizontalAlignment.CENTER
    var verticalAlignment: VerticalAlignment = VerticalAlignment.CENTER

    override fun create(excelElement: ExcelElement, workbook: Workbook, onSheet: Int?) {
        val sheet = workbook.getSheetAt(onSheet ?: 0)
        data?.forEachIndexed { rowIdx, rowData ->
            val row = sheet.createRow(rowIdx.plus(1))
            rowData.forEachIndexed { index, cellData ->
                val cell = row.createCell(index)
                cell?.setCellValue(HSSFRichTextString(cellData.toString()))
                val style = sheet?.workbook?.createCellStyle()
                style?.alignment = horizontalAlignment
                style?.verticalAlignment = verticalAlignment
                cell?.cellStyle = style
            }
        }
    }

}