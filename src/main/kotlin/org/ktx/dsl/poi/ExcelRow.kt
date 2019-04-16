package org.ktx.dsl.poi

import org.apache.poi.hssf.usermodel.HSSFRichTextString
import org.apache.poi.ss.usermodel.*

open class ExcelRow {

    var data: List<Any>? = null
    var horizontalAlignment: HorizontalAlignment = HorizontalAlignment.CENTER
    var verticalAlignment: VerticalAlignment = VerticalAlignment.CENTER

    private val cellList: MutableList<ExcelCell> = mutableListOf()
    private var insertDataBefore: Boolean = false

    fun cell(init: ExcelCell.() -> Unit): ExcelCell {
        val cell = ExcelCell()
        cell.init()
        if (!data.isNullOrEmpty()) {
            insertDataBefore = true
        }
        cellList.add(cell)
        return cell
    }

    fun cells(init: ExcelCells.() -> Unit): ExcelCells {
        val excelCells = ExcelCells()
        excelCells.init()
        excelCells.cells?.forEach {
            val cell = ExcelCell()
            cell.style = excelCells.commonStyle
            cell.value = it.toString()
            if (!data.isNullOrEmpty()) {
                insertDataBefore = true
            }
            cellList.add(cell)
        }
        if (excelCells.cellsWithStyle.isNullOrEmpty().not()) {
            if (!cellList.containsAll(excelCells.cellsWithStyle!!)) {
                cellList.addAll(excelCells.cellsWithStyle!!)
            }
        }
        return excelCells
    }

    fun create(excelElement: ExcelElement, workbook: Workbook, onSheet: Int?, rowIndex: Int?) {
        val sheet = workbook.getSheetAt(onSheet ?: 0)

        val internalDataList = mutableListOf<ExcelCell>()
        data?.forEach {
            val cell = ExcelCell()
            cell.value = it.toString()
            internalDataList.add(cell)
        }

        if (insertDataBefore) {
            cellList.addAll(0, internalDataList)
        } else {
            cellList.addAll(internalDataList)
        }

        if (excelElement is ExcelSheet) {
            val row = sheet.createRow(rowIndex ?: 0)
            cellList.forEachIndexed { index, data ->
                data.create(workbook, row, index)
            }
        }
    }

    private fun defaultStyle(sheet: Sheet): CellStyle {
        val style = sheet.workbook.createCellStyle()
        style.alignment = horizontalAlignment
        style.verticalAlignment = verticalAlignment
        return style
    }
}