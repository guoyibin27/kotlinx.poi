package org.ktx.dsl.poi

import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.util.CellRangeAddress

class Main

fun main() {

    val workbook = workbook {
        fileName = "test.xls"
        filePath = "d:\\questionlib"
        sheet {
            autoSizeColumn = arrayOf(2, 3, 4)
            row {
                data = listOf("cell11", "cell12", "cell13", "cell14", "cell15", "cell16")
            }
            titleRow {
                data = listOf("标题11", "标题12", "标题13", "标题14", "标题15", "标题16")
            }
            row {
                cell {
                    value = "add row 11"
                }
                cell {
                    value = "add row 12"
                }
                cell {
                    value = "add row 13"
                }
                cell {
                    value = "add row 14"
                }
                cell {
                    value = "add row 15"
                    style {
                        font {
                            bold = true
                            underline = HSSFFont.U_DOUBLE_ACCOUNTING
                            color = HSSFColor.HSSFColorPredefined.RED.index
                        }
                    }
                }
                cell {
                    value = "add row 16"
                    style {
                        borderLeft = BorderStyle.HAIR
                        borderRight = BorderStyle.DASH_DOT
                        borderTop = BorderStyle.DOTTED
                        rightBorderColor = HSSFColor.HSSFColorPredefined.RED.index
                        bottomBorderColor = HSSFColor.HSSFColorPredefined.RED.index
                        shrinkToFit = true
                        hidden = true
                        font {
                            bold = true
                            underline = HSSFFont.U_DOUBLE_ACCOUNTING
                            color = HSSFColor.HSSFColorPredefined.YELLOW.index
                        }
                    }
                }
            }
            row {
                cells {
                    cells = listOf("new cell1", "new cell1", "new cell1")
                    commonStyle = commonStyle {
                        font {
                            bold = true
                            underline = HSSFFont.U_DOUBLE_ACCOUNTING
                            color = HSSFColor.HSSFColorPredefined.CORAL.index
                        }
                    }
                    cellsWithStyle = listOf(
                        cell {
                            value = "newcell1"
                            style {
                                font {
                                    bold = true
                                    underline = HSSFFont.U_DOUBLE_ACCOUNTING
                                    color = HSSFColor.HSSFColorPredefined.BLUE.index
                                }
                            }
                        },
                        cell {
                            value = "newcell2"
                            style {
                                font {
                                    bold = true
                                    underline = HSSFFont.U_DOUBLE_ACCOUNTING
                                    color = HSSFColor.HSSFColorPredefined.PINK.index
                                }
                            }
                        }, cell {
                            value = "newcell3"
                            style {
                                font {
                                    bold = true
                                    underline = HSSFFont.U_DOUBLE_ACCOUNTING
                                    color = HSSFColor.HSSFColorPredefined.AQUA.index
                                }
                            }
                        }
                    )
                }
            }
            titleWithRow {
                titles = listOf(
                    listOf("标题21", "标题22", "标题23", "标题24", "标题25", "标题26"),
                    listOf("标题31", "标题32", "标题33", "标题34", "标题35", "标题36")
                )
                rows = listOf(
                    listOf("cell3", "cell3", "cell13", "cell14", "cell15", "cell16"),
                    listOf("cell4", "cell4", "cell4", "cell4", "cell4", "cell4"),
                    listOf("cell5", "cell5", "cell5", "cell5", "cell5", "cell5")
                )
            }

            rows {
                data = listOf(
                    listOf("cell3", "cell3", "cell13", "cell14", "cell15", "cell16"),
                    listOf("cell4", "cell4", "cell4", "cell4", "cell4", "cell4"),
                    listOf("cell5", "cell5", "cell5", "cell5", "cell5", "cell5")
                )
            }
        }
    }
}