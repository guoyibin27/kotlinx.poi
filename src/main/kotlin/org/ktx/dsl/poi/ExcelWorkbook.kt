package org.ktx.dsl.poi

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import java.io.FileOutputStream


class ExcelWorkbook : ExcelElement {

    var filePath: String? = null
    var fileName: String? = null

    private val sheetList: MutableList<ExcelSheet> by lazy { mutableListOf<ExcelSheet>() }
    private val workbook: Workbook by lazy { HSSFWorkbook() }

    fun sheet(init: ExcelSheet.(ExcelWorkbook) -> Unit): ExcelSheet {
        val excelSheet = ExcelSheet()
        sheetList.add(excelSheet)
        excelSheet.init(this)
        excelSheet.create(excelSheet, workbook, sheetList.count().minus(1))
        return excelSheet
    }

    override fun create(excelElement: ExcelElement, workbook: Workbook, onSheet: Int?) {
    }

    fun save() {
        assert(filePath.isNullOrEmpty()) { "文件路径不存在" }
        assert(fileName.isNullOrEmpty()) { "文件名称不存在" }
        var stream: FileOutputStream? = null
        try {
            stream = FileOutputStream("$filePath/$fileName")
            workbook.write(stream)
        } finally {
            stream?.flush()
            stream?.close()
        }
    }
}

fun workbook(init: ExcelWorkbook.(ExcelWorkbook) -> Unit): ExcelWorkbook {
    val workbook = ExcelWorkbook()
    workbook.init(workbook)
    workbook.save()
    return workbook
}

