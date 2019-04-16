package org.ktx.dsl.poi

import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.Color

class ExcelCellFont {

    var name: String? = null
    var heightInPoint: Short? = null
    var italic: Boolean = false
    var strikeout: Boolean = false
    var color: Short = HSSFColor.HSSFColorPredefined.BLACK.index
    var typeOffset: Short = HSSFFont.SS_NONE
    var underline: Byte = HSSFFont.U_NONE
    var charset: Byte = HSSFFont.DEFAULT_CHARSET
    var bold: Boolean = false
}