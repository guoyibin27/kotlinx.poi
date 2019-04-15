package org.ktx.dsl.poi

class Main

fun main() {

    val workbook = workbook {
        fileName = "test.xls"
        filePath = "d:\\questionlib"
        sheet {
            titleRow {
                titles = listOf("标题1", "标题2", "标题3", "标题4", "标题5", "标题6")
            }
            row {
                data = listOf(
                    listOf("cell11", "cell12", "cell13", "cell14", "cell15", "cell16"),
                    listOf("cell21", "cell22", "cell23", "cell24", "cell25", "cell26"),
                    listOf("cell31", "cell32", "cell33", "cell34", "cell35", "cell36")
                )
            }
        }
    }
}