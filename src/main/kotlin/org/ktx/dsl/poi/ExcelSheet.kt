package org.ktx.dsl.poi

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress


class ExcelSheet : ExcelElement {
    var name: String? = null
    var mergedRegion: CellRangeAddress? = null
    var autoSizeColumn: Array<Int>? = null
    private var workbook: Workbook? = null
    private var sheetIndex: Int = 0
    private var sheet: Sheet? = null
    private var rows: MutableList<ExcelRow> = mutableListOf()
    private var titleRows: MutableList<ExcelTitleRow> = mutableListOf()

    fun titleRow(init: ExcelTitleRow.() -> Unit): ExcelTitleRow {
        val titleRow = ExcelTitleRow()
        titleRow.init()
        titleRows.add(titleRow)
        return titleRow
    }

    fun row(init: ExcelRow.() -> Unit): ExcelRow {
        val row = ExcelRow()
        row.init()
        rows.add(row)
        return row
    }

    fun rows(init: ExcelRows.() -> Unit): ExcelRows {
        val rowList = ExcelRows()
        rowList.init()
        rowList.data?.forEach {
            val row = ExcelRow()
            row.data = it
            rows.add(row)
        }
        return rowList
    }

    fun titleRows(init: ExcelTitleRows.() -> Unit): ExcelTitleRows {
        val titles = ExcelTitleRows()
        titles.init()
        titles.titles?.forEach {
            val titleRow = ExcelTitleRow()
            titleRow.data = it
            titleRows.add(titleRow)
        }
        return titles
    }

    fun titleWithRow(init: ExcelTitleWithRow.() -> Unit): ExcelTitleWithRow {
        val excelTitleWithRow = ExcelTitleWithRow()
        excelTitleWithRow.init()

        excelTitleWithRow.titles?.forEach {
            val titleRow = ExcelTitleRow()
            titleRow.data = it
            titleRows.add(titleRow)
        }

        excelTitleWithRow.rows?.forEach {
            val row = ExcelRow()
            row.data = it
            rows.add(row)
        }
        return excelTitleWithRow
    }

    override fun create(excelElement: ExcelElement, workbook: Workbook, onSheet: Int?) {
        this.sheetIndex = onSheet ?: 0
        this.workbook = workbook
        sheet = if (name.isNullOrEmpty()) {
            workbook.createSheet()
        } else {
            workbook.createSheet(name)
        }
        if (mergedRegion != null) {
            sheet!!.addMergedRegion(mergedRegion)
        }
        titleRows.forEachIndexed { idx, it ->
            it.create(this, workbook, sheetIndex, idx)
        }
        rows.forEachIndexed { idx, it ->
            it.create(this, workbook, sheetIndex, titleRows.count().plus(idx))
        }
        autoSizeColumn?.forEach {
            sheet!!.autoSizeColumn(it, mergedRegion != null)
        }
    }

}