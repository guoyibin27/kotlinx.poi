package org.ktx.dsl.poi

import org.apache.poi.ss.usermodel.Workbook

@ExcelMarker
interface ExcelRowElement : ExcelElement {

    fun create(excelElement: ExcelElement, workbook: Workbook, onSheet: Int? = null, rowIndex: Int? = null) {

    }
}