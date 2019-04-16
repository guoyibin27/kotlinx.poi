package org.ktx.dsl.poi

class ExcelCells : ExcelElement {

    var cells: List<Any>? = null

    var cellsWithStyle: List<ExcelCell>? = null
    var commonStyle: ExcelCellStyle? = null

    fun commonStyle(init: ExcelCellStyle.() -> Unit): ExcelCellStyle {
        commonStyle = ExcelCellStyle()
        commonStyle!!.init()
        return commonStyle!!
    }

}