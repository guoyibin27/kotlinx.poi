package org.ktx.dsl.poi

import org.apache.poi.ss.usermodel.*

class ExcelCellStyle {

    /**
     * set the data format (must be a valid format). Built in formats are defined at [BuiltinFormats].
     * @see DataFormat
     */
    var dataFormat: Short? = null

    /**
     * set the cell's using this style to be hidden
     * @param hidden - whether the cell using this style should be hidden
     */
    var hidden: Boolean = false


    /**
     * set the cell's using this style to be locked
     * @param locked - whether the cell using this style should be locked
     */
    var locked: Boolean = false

    /**
     * Turn on or off "Quote Prefix" or "123 Prefix" for the style,
     * which is used to tell Excel that the thing which looks like
     * a number or a formula shouldn't be treated as on.
     * Turning this on is somewhat (but not completely, see [IgnoredErrorType])
     * like prefixing the cell value with a ' in Excel
     */
    var quotePrefixed: Boolean = false


    /**
     * set the type of horizontal alignment for the cell
     * @param align - the type of alignment
     */
    var alignment: HorizontalAlignment? = null


    /**
     * Set whether the text should be wrapped.
     * Setting this flag to `true` make all content visible
     * within a cell by displaying it on multiple lines
     *
     * @param wrapped  wrap text or not
     */
    var wrapText: Boolean = false

    /**
     * set the type of vertical alignment for the cell
     * @param align the type of alignment
     */
    var verticalAlignment: VerticalAlignment? = null

    /**
     * set the degree of rotation for the text in the cell.
     *
     * Note: HSSF uses values from -90 to 90 degrees, whereas XSSF
     * uses values from 0 to 180 degrees. The implementations of this method will map between these two value-ranges
     * accordingly, however the corresponding getter is returning values in the range mandated by the current type
     * of Excel file-format that this CellStyle is applied to.
     *
     * @param rotation degrees (see note above)
     */
    var rotation: Short? = null

    /**
     * set the number of spaces to indent the text in the cell
     * @param indent - number of spaces
     */
    var indention: Short? = 0

    /**
     * set the type of border to use for the left border of the cell
     * @param border type
     * @since POI 3.15
     */
    var borderLeft: BorderStyle = BorderStyle.NONE

    /**
     * set the type of border to use for the right border of the cell
     * @param border type
     * @since POI 3.15
     */
    var borderRight: BorderStyle = BorderStyle.NONE

    /**
     * set the type of border to use for the top border of the cell
     * @param border type
     * @since POI 3.15
     */
    var borderTop: BorderStyle = BorderStyle.NONE


    /**
     * set the type of border to use for the bottom border of the cell
     * @param border type
     * @since POI 3.15
     */
    var borderBottom: BorderStyle = BorderStyle.NONE


    /**
     * set the color to use for the left border
     * @param color The index of the color definition
     */
    var leftBorderColor: Short? = null


    /**
     * set the color to use for the right border
     * @param color The index of the color definition
     */
    var rightBorderColor: Short? = null


    /**
     * set the color to use for the top border
     * @param color The index of the color definition
     */
    var topBorderColor: Short? = null

    /**
     * set the color to use for the bottom border
     * @param color The index of the color definition
     */
    var bottomBorderColor: Short? = null

    /**
     * setting to one fills the cell with the foreground color... No idea about
     * other values
     *
     * @param fp  fill pattern (set to [FillPatternType.SOLID_FOREGROUND] to fill w/foreground color)
     * @since POI 3.15 beta 3
     */
    var fillPattern: FillPatternType? = null

    /**
     * set the background fill color.
     *
     * @param bg  color
     */
    var fillBackgroundColor: Short? = null

    /**
     * Gets the color object representing the current
     * background fill, resolving indexes using
     * the supplied workbook.
     * This will work for both indexed and rgb
     * defined colors.
     */
    var fillBackgroundColorColor: Short? = null


    /**
     * set the foreground fill color
     * *Note: Ensure Foreground color is set prior to background color.*
     * @param bg  color
     */
    var fillForegroundColor: Short? = null

    /**
     * Controls if the Cell should be auto-sized
     * to shrink to fit if the text is too long
     */
    var shrinkToFit: Boolean = false

    private var font: ExcelCellFont? = null

    fun font(init: ExcelCellFont.() -> Unit): ExcelCellFont {
        font = ExcelCellFont()
        font!!.init()
        return font!!
    }

    fun getCellStyle(workbook: Workbook, style: CellStyle) {
        if (this.alignment != null) {
            style.alignment = this.alignment
        }
        if (this.verticalAlignment != null) {
            style.verticalAlignment = this.verticalAlignment
        }
        style.borderBottom = this.borderBottom
        style.borderLeft = this.borderLeft
        style.borderRight = this.borderRight
        style.borderTop = this.borderTop
        if (this.dataFormat != null) {
            style.dataFormat = this.dataFormat!!
        }
        style.hidden = this.hidden
        style.locked = this.locked
        style.quotePrefixed = this.quotePrefixed
        style.wrapText = this.wrapText
        if (this.rotation != null) {
            style.rotation = this.rotation!!
        }
        if (this.indention != null) {
            style.indention = this.indention!!
        }
        if (this.leftBorderColor != null) {
            style.leftBorderColor = this.leftBorderColor!!
        }
        if (this.rightBorderColor != null) {
            style.rightBorderColor = this.rightBorderColor!!
        }
        if (this.topBorderColor != null) {
            style.topBorderColor = this.topBorderColor!!
        }
        if (this.bottomBorderColor != null) {
            style.bottomBorderColor = this.bottomBorderColor!!
        }
        if (this.fillPattern != null) {
            style.fillPattern = this.fillPattern!!
        }
        if (this.fillBackgroundColor != null) {
            style.fillBackgroundColor = this.fillBackgroundColor!!
        }
        if (this.fillBackgroundColorColor != null) {
            style.fillBackgroundColor = this.fillBackgroundColorColor!!
        }
        if (this.fillForegroundColor != null) {
            style.fillForegroundColor = this.fillForegroundColor!!
        }
        style.shrinkToFit = this.shrinkToFit
        getFont(workbook, style)
    }

    private fun getFont(workbook: Workbook, style: CellStyle) {
        if (this.font == null) return
        val newFont = workbook.createFont()
        newFont.bold = this.font!!.bold
        if (this.font!!.name != null) {
            newFont.fontName = this.font!!.name
        }
        if (this.font!!.heightInPoint != null) {
            newFont.fontHeightInPoints = this.font!!.heightInPoint!!
        }
        newFont.italic = this.font!!.italic
        newFont.strikeout = this.font!!.strikeout
        newFont.color = this.font!!.color
        newFont.typeOffset = this.font!!.typeOffset
        newFont.underline = this.font!!.underline
        newFont.setCharSet(this.font!!.charset)
        style.setFont(newFont)
    }
}