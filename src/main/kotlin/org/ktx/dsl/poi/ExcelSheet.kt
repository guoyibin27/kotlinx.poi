package org.ktx.dsl.poi

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook


class ExcelSheet : ExcelElement {
    var name: String? = null
    private var workbook: Workbook? = null
    private var sheetIndex: Int = 0
    private var sheet: Sheet? = null
    fun titleRow(init: ExcelTitleRow.() -> Unit): ExcelTitleRow {
        if (sheet == null) {
            sheet = if (name.isNullOrEmpty()) {
                workbook?.createSheet()
            } else {
                workbook?.createSheet(name)
            }
        }
        val titleRow = ExcelTitleRow()
        titleRow.init()
        titleRow.create(this, workbook!!, sheetIndex)
        return titleRow
    }

    fun row(init: ExcelRow.() -> Unit): ExcelRow {
        if (sheet == null) {
            sheet = if (name.isNullOrEmpty()) {
                workbook?.createSheet()
            } else {
                workbook?.createSheet(name)
            }
        }
        val row = ExcelRow()
        row.init()
        row.create(this, workbook!!, sheetIndex)
        return row
    }

    override fun create(excelElement: ExcelElement, workbook: Workbook, onSheet: Int?) {
        this.sheetIndex = onSheet ?: 0
        this.workbook = workbook
    }
}