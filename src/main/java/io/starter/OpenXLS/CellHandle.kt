/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with OpenXLS.  If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package io.starter.OpenXLS

import io.starter.formats.XLS.Font
import io.starter.formats.XLS.*
import io.starter.formats.XLS.charts.Ai
import io.starter.formats.XLS.formulas.Ptg
import io.starter.formats.XLS.formulas.PtgRef
import io.starter.formats.cellformat.CellFormatFactory
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.json.JSONException
import org.json.JSONObject

import java.awt.Color
import java.awt.font.TextAttribute
import java.io.Serializable
import java.util.*

import io.starter.OpenXLS.JSONConstants.*

/**
 * The CellHandle provides a handle to an XLS Cell and its values. <br></br>
 * Use the CellHandle to work with individual Cells in an XLS file. <br></br>
 * To instantiate a CellHandle, you must first have a valid WorkSheetHandle,
 * which in turn requires a valid WorkBookHandle. <br></br>
 * <br></br>
 * for example: <br></br>
 * <br></br>
 * <blockquote> WorkBookHandle book = newWorkBookHandle("testxls.xls");<br></br>
 * WorkSheetHandle sheet1 = book.getWorkSheet("Sheet1");<br></br>
 * CellHandlecell = sheet1.getCell("B22");<br></br>
 * <br></br>
 * <br></br>
</blockquote> *  With a CellHandle you can:<br></br>
 * <br></br>
 * - get the value of a Cell<br></br>
 * - set the value of a Cell<br></br>
 * - change the color, font, and background formatting of a Cell<br></br>
 * - change the formatting pattern of a Cell<br></br>
 * - change the value of a Cell<br></br>
 * - get a handle to any Formula for this Cell<br></br>
 * <br></br>
 *
 * [Starter Inc.](http://starter.io)
 *
 * @see WorkBookHandle
 *
 * @see WorkSheetHandle
 *
 * @see FormulaHandle
 *
 * @see CellNotFoundException
 *
 * @see CellTypeMismatchException
 */
class CellHandle : Cell, Serializable, Handle, Comparable<CellHandle> {

    /**
     * returns the WorkBookHandle for this Cell
     *
     * @return WorkBook the book
     */
    @Transient
    val workBook: WorkBook? = null
    /**
     * get the WorkSheetHandle for this Cell
     *
     * @return the WorkSheetHandle for this Cell
     */
    /**
     * set the WorkSheetHandle for this Cell
     *
     * @param WorkSheetHandle handle - the new worksheet for this Cell
     * @see WorkSheetHandle
     */
    @Transient // This is redundant, already done in WSH.getCell().
    // if (wsh!=null) //20080616 KSC
    // wsh.cellhandles.put(this.getCellAddress(), this);
    var workSheetHandle: WorkSheetHandle? = null
    private var mycol: ColHandle? = null
    private var myrow: RowHandle? = null
    private var formatter: FormatHandle? = null
    // reusing or creating new xfs is handled in FormatHandle/cloneXf and
    // updateXf
    // boolean useExistingXF = !false;
    private val DEBUG = false
    /**
     * returns the underlying BIFF8 record for the Cell <br></br>
     * NOTE: the underlying record is not a part of the public API and may change at
     * any time.
     *
     * @return Returns the underlying biff record.
     */
    /**
     * sets the underlying BIFF8 record for the Cell <br></br>
     * NOTE: the underlying record is not a part of the public API and may change at
     * any time.
     *
     * @param XLSRecord rec - The BIFF record to set.
     */
    var record: XLSRecord? = null

    /**
     * Get the default "empty" data value for this CellHandle
     *
     * @return Object a default empty value corresponding to the cell type <br></br>
     * such as 0.0 for TYPE_DOUBLE or an empty String for TYPE_BLANK
     */
    val defaultVal: Any?
        get() = record!!.defaultVal

    /**
     * Gets the number format pattern for this cell, if set. For more information on
     * number format patterns see
     * [Microsoft KB264372](http://support.microsoft.com/kb/264372).
     *
     * @return the Excel number format pattern for this cell or `null` if
     * none is applied
     */
    /**
     * Sets the number format pattern for this cell. All Excel built-in number
     * formats are supported. Custom formats will not be applied by OpenXLS (e.g.
     * the [.getFormattedStringVal] method) but they will be written correctly
     * to the output file. For more information on number format patterns see
     * [Microsoft KB264372](http://support.microsoft.com/kb/264372).
     *
     * @param pat the Excel number format pattern to apply
     * @see FormatConstantsImpl.getBuiltinFormats
     */
    var formatPattern: String?
        get() = if (this.font == null) "" else record!!.formatPattern
        set(pat) {
            setFormatHandle()
            formatter!!.formatPattern = pat
        }

    /**
     * Returns whether this Cell has Date formatting applied. <br></br>
     * NOTE: This does not guarantee that the value is a valid date.
     *
     * @return boolean true if this Cell has a Date Format applied
     */
    override val isDate: Boolean
        get() {
            if (record!!.myxf == null)
                return false
            if (record!!.isString)
                return false
            if (record!!.isBoolean)
                return false
            return if (record!!.isBlank) false else record!!.myxf!!.isDatePattern
        }

    /**
     * Returns whether this cell is a formula.
     */
    val isFormula: Boolean
        get() = record!!.isFormula

    /**
     * Returns whether the formula for the Cell is hidden
     *
     * @return boolean true if formula is hidden
     */
    /**
     * Hides or shows the formula for this Cell, if present
     *
     * @param boolean hidden - setting whether to hide or show formulas for this Cell
     */
    // create a new xf for this
    // this causes formats to be lost this.useExistingXF = false;
    var isFormulaHidden: Boolean
        get() = this.formatHandle!!.isFormulaHidden
        set(hidden) {
            formatHandle!!.isFormulaHidden = hidden
        }

    /**
     * returns whether this Cell is locked for editing
     *
     * @return boolean true if the cell is locked
     */
    /**
     * locks or unlocks this Cell for editing
     *
     * @param boolean locked - true if Cell should be locked, false otherwise
     */
    // create a new xf for this
    // this causes formats to be lost this.useExistingXF = false;
    var isLocked: Boolean
        get() = this.formatHandle!!.isLocked
        set(locked) {
            formatHandle!!.isLocked = locked
        }

    /**
     * Returns whether this is a blank cell.
     *
     * @return true if this cell is blank
     */
    val isBlank: Boolean
        get() = this.record!!.isBlank

    /**
     * Returns whether this Cell has a Currency format applied.
     *
     * @return boolean true if this Cell has a Currency format applied
     */
    val isCurrency: Boolean
        get() = if (record!!.myxf == null) false else record!!.myxf!!.isCurrencyPattern

    /**
     * Returns whether this Cell has a numeric value
     *
     * @return boolean true if this Cell contains a numeric value
     */
    val isNumber: Boolean
        get() = this.record!!.isNumber

    /**
     * get the weight (boldness) of the Font used by this Cell.
     *
     * @return int Font weight range between 100-1000
     */
    /**
     * change the weight (boldness) of the Font used by this Cell. <br></br>
     * Some examples: <br></br>
     * a weight of 200 is normal <br></br>
     * a weight of 700 is bold <br></br>
     *
     * @param int wt - Font weight range between 100-1000
     */
    var fontWeight: Int
        get() = if (this.font == null) FormatHandle.DEFAULT_FONT_WEIGHT else record!!.font!!.fontWeight
        set(wt) {
            setFormatHandle()
            formatter!!.fontWeight = wt
        }

    /**
     * get the size in points of the Font used by this Cell
     *
     * @return int Font size in Points.
     */
    /**
     * change the size (in points) of the Font used by this Cell and all other Cells
     * sharing this FormatId. <br></br>
     * NOTE: To add an entirely new Font use the setFont(String fn, int typ, int sz)
     * method instead.
     *
     * @param int sz - Font size in Points.
     */
    // excel size is 1/20 pt.
    var fontSize: Int
        get() = if (this.font == null) FormatHandle.DEFAULT_FONT_SIZE else record!!.font!!.fontHeight / 20
        set(sz) {
            var sz = sz
            setFormatHandle()
            sz *= 20
            formatter!!.fontHeight = sz
        }

    /**
     * get the Color of the Font used by this Cell. <br></br>
     *
     * @return int Excel color constant for Font color
     * @see FormatHandle.COLOR constants
     */
    /**
     * Set the color of the Font used by this Cell.
     *
     * @param java .Awt.Color col - color of the font
     */
    // handle white on white text issue
    // return black
    // black on black
    var fontColor: Color
        get() {
            if (this.font == null)
                return FormatHandle.Black
            val clidx = this.font!!.color
            val x = record!!.xfRec
            val clidb = x!!.backgroundColor.toInt()
            if (clidx == 64 && clidb == 64) {
                return FormatHandle.Black
            }
            return if (clidx < this.workBook!!.colorTable.size) {
                workBook.colorTable[clidx]
            } else Color.black
        }
        set(col) {
            setFormatHandle()
            formatter!!.fontColor = col
        }

    /**
     * get the Color of the Cell Foreground of this Cell and all other cells sharing
     * this FormatId. <br></br>
     * <br></br>
     * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
     * Background color is the PATTERN color for all patterns not equal to
     * PATTERN_SOLID <br></br>
     * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
     * Background Color is 64 (white).
     *
     * @return java.awt.Color of the Font
     */
    val foregroundColor: Color
        get() {
            if (this.record!!.xfRec != null) {
                val clidx = record!!.xfRec!!.foregroundColor.toInt()
                if (clidx < this.workBook!!.colorTable.size) {
                    return workBook.colorTable[clidx]
                }
            }
            return Color.black
        }

    /**
     * get the Color of the Cell Background by this Cell and all other cells sharing
     * this FormatId. <br></br>
     * <br></br>
     * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
     * Background color is the PATTERN color for all patterns not equal to
     * PATTERN_SOLID <br></br>
     * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
     * Background Color is 64 (white).
     *
     * @return java.awt.Color for the Font
     */
    val backgroundColor: Color
        get() {
            if (this.record!!.xfRec != null) {
                val x = record!!.xfRec
                val clidx = x!!.backgroundColor.toInt()
                if (clidx < this.workBook!!.colorTable.size) {
                    return workBook.colorTable[clidx]
                }
            }
            return Color.white
        }

    /**
     * Return the cell background color i.e. the color if a pattern is set, or white
     * if none
     *
     * @return java.awt.Color cell background color
     */
    val cellBackgroundColor: Color
        get() {
            setFormatHandle()
            val clidx = formatter!!.cellBackgroundColor
            return if (clidx < this.workBook!!.workBook.colorTable.size) {
                workBook.colorTable[clidx]
            } else Color.white
        }

    /**
     * get the Background Pattern for this cell <br></br>
     *
     * @return int Excel pattern constant
     * @see FormatHandle.PATTERN constants
     */
    /**
     * set the Background Pattern for this Cell. <br></br>
     *
     * @param int t - Excel pattern constant
     * @see FormatHandle.PATTERN constants
     */
    var backgroundPattern: Int
        get() = this.record!!.xfRec!!.fillPattern
        set(t) {
            setFormatHandle()
            formatter!!.setPattern(t)
        }

    /**
     * get the fill pattern for this cell <br></br>
     * Same as getBackgroundPattern <br></br>
     *
     * @return int Excel fill pattern constant
     * @see FormatHandle.PATTERN constants
     */
    val fillPattern: Int
        get() {
            setFormatHandle()
            return formatter!!.fillPattern
        }

    /**
     * get the Font face used by this Cell.
     *
     * @return String the system name of the font for this Cell
     */
    /**
     * set the Font face used by this Cell.
     *
     * @param String fn - system name of the font
     */
    var fontFace: String?
        get() = if (this.font == null) FormatHandle.DEFAULT_FONT_FACE else record!!.font!!.fontName
        set(fn) {
            setFormatHandle()
            formatter!!.fontName = fn
        }

    /**
     * Get the OpenXLS Font for this Cell, which roughly matches the functionality
     * of the java.awt.Font class. <br></br>
     * Due to awt problems on console systems, converting the OpenXLS font to a GUI
     * font is up to you. <br></br>
     *
     * @return OpenXLS font for Cell
     */
    val font: Font?
        get() = record!!.font

    /**
     * convert the OpenXLS font used by this Cell to java.awt.Font
     *
     * @return java.awt.Font for this cell
     */
    // back to default
    // Excel seems to display 1 point larger than Java
    // Logger.logInfo("Displaying font:" + fface);
    // implement underlines
    // Ahhh, much cooler fonts here!
    // TODO: Interpret other weights- LIGHT, DEMI_LIGHT, DEMI_BOLD, etc.
    val awtFont: java.awt.Font
        get() {
            var fface: String? = "Arial"
            try {
                fface = this.fontFace
            } catch (e: Exception) {
            }

            var sz = 12
            try {
                sz = this.fontSize
            } catch (e: Exception) {
                sz = 12
            }

            sz += 4
            val ftmap = HashMap<TextAttribute, Serializable>()
            try {
                if (this.isUnderlined) {
                    ftmap[java.awt.font.TextAttribute.UNDERLINE] = java.awt.font.TextAttribute.UNDERLINE
                }
            } catch (e: Exception) {
            }

            ftmap[java.awt.font.TextAttribute.FAMILY] = fface
            val fx = this.fontWeight.toFloat()
            ftmap[java.awt.font.TextAttribute.SIZE] = sz
            if (fx == FormatHandle.BOLD.toFloat())
                ftmap[java.awt.font.TextAttribute.WEIGHT] = java.awt.font.TextAttribute.WEIGHT_BOLD
            else
                ftmap[java.awt.font.TextAttribute.WEIGHT] = java.awt.font.TextAttribute.WEIGHT_REGULAR
            return java.awt.Font(ftmap)
        }

    /**
     * Get a CommentHandle to the note attached to this cell
     */
    // this needs significant cleanup. We should not have to iterate notes
    val comment: CommentHandle
        @Throws(DocumentObjectNotFoundException::class)
        get() {
            val notes = record!!.sheet!!.notes
            for (i in notes.indices) {
                val n = notes[i] as Note
                if (n.cellAddressWithSheet == this.cellAddressWithSheet) {
                    return CommentHandle(n)
                }
            }
            throw DocumentObjectNotFoundException("Note record not found at " + this.cellAddressWithSheet)
        }

    /**
     * returns whether the Font for this Cell is underlined
     *
     * @return boolean true if Font is underlined
     */
    /**
     * Set whether the Font for this Cell is underlined
     *
     * @param boolean b - true if the Font for this Cell should be underlined (single
     * underline only)
     */
    var isUnderlined: Boolean
        get() = if (this.font == null) false else this.font!!.underlineStyle != 0x0
        set(isUnderlined) {
            setFormatHandle()
            if (isUnderlined)
                this.font!!.setUnderlineStyle(Font.STYLE_UNDERLINE_SINGLE)
            else
                this.font!!.setUnderlineStyle(Font.STYLE_UNDERLINE_NONE)
        }

    /**
     * Returns any other Cells merged with this one, or null if this Cell is not a
     * part of a merged range. <br></br>
     * Adding and/or removing Cells from this CellRange will merge or unmerge the
     * Cells.
     *
     * @return CellRange object containing all Cells in this Cell's merged range.
     */
    val mergedCellRange: CellRange?
        get() = record!!.mergeRange

    /**
     * Returns if the Cell is the parent (cell containing display value) of a merged
     * cell range.
     *
     * @return boolean true if this cell is a merge parent cell
     */
    val isMergeParent: Boolean
        get() {
            try {
                val cr = record!!.mergeRange ?: return false
                val i = cr.rangeCoords
                if (this.rowNum + 1 == i[0] && this.colNum == i[1])
                    return true
            } catch (e: Exception) {
                return false
            }

            return false

        }

    /**
     * Returns the ColHandle for the Cell.
     *
     * @return ColHandle for the Cell
     */
    // can't happen, the column has to exist because we're in it
    val col: ColHandle
        get() {
            try {
                if (mycol == null)
                    mycol = workSheetHandle!!.getCol(this.colNum)
            } catch (ex: ColumnNotFoundException) {
                throw RuntimeException(ex)
            }

            return this.mycol
        }

    /**
     * Returns the RowHandle for the Cell.
     *
     * @return RowHandle representing the Row for the Cell
     */
    val row: RowHandle
        get() {
            if (myrow == null)
                myrow = RowHandle(record!!.row, this.workSheetHandle!!)
            return this.myrow
        }

    /**
     * Returns the value of this Cell in the native underlying data type.
     *
     *
     * Formula cells will return the calculated value of the formula in the
     * calculated data type.
     *
     *
     * Use 'getStringVal()' to return a String regardless of underlying value type.
     *
     * @return Object value for this Cell
     */
    /**
     * Set the val of this Cell to an object <br></br>
     * <br></br>
     * The object may be one of type: <br></br>
     * String, Float, Integer, Double, Long, Boolean, java.sql.Date, or null <br></br>
     * <br></br>
     * To set a Cell to a formula, obj should be a string begining with "=" <br></br>
     * <br></br>
     * To set a Cell to an array formula, obj should be a string begining with "{="
     * <br></br>
     * If you wish to put a line break in a string value, use the newline "\n"
     * character. Note this will not function unless you also apply a format handle
     * to the cell with WrapText=true
     *
     * @param Object obj - the object to set the value of the Cell to
     * @throws CellTypeMismatchException
     */

    override// blow out cache
    // Formula or array string
            /* 20070212 KSC: FunctionNotSupported */// suppress these -- cell has been changed
    // NOT a CTMME always
    var `val`: Any?
        get() = FormulaHandle.sanitizeValue(record!!.internalVal)
        @Throws(CellTypeMismatchException::class)
        set(obj) {
            if (this.workBook!!.formulaCalculationMode != WorkBook.CALCULATE_EXPLICIT)
                this.clearAffectedCells()

            if (obj is java.sql.Date) {
                this.setVal(obj as java.sql.Date?, null)
                return
            }
            if (obj is String) {
                val formstr = obj as String?
                if (formstr!!.indexOf("=") == 0 || formstr.startsWith("{=")) {
                    try {
                        this.setFormula(formstr)
                        return
                    } catch (a: Exception) {
                        Logger.logWarn(
                                "CellHandle.setVal() failed.  Setting Formula to $obj failed: $a")
                    }

                }
            }
            try {
                setBiffRecValue(obj)
            } catch (fnse: FunctionNotSupportedException) {
            } catch (e: Exception) {

                throw CellTypeMismatchException(e.toString())
            }

        }

    /**/

    /**
     * returns the java Type string of the Cell <br></br>
     * One of:
     *  * "String"
     *  * "Float"
     *  * "Integer"
     *  * "Formula"
     *  * "Double"
     *
     * @return String java data type
     */
    val cellTypeName: String
        get() {
            var typename = "Object"
            val tp = cellType
            when (tp) {
                XLSConstants.TYPE_BLANK -> typename = "String"
                XLSConstants.TYPE_STRING -> typename = "String"
                XLSConstants.TYPE_FP -> typename = "Float"
                XLSConstants.TYPE_INT -> typename = "Integer"
                XLSConstants.TYPE_FORMULA -> typename = "Formula"
                XLSConstants.TYPE_DOUBLE -> typename = "Double"
            }
            return typename
        }

    /**
     * Returns the type of the Cell as an int <br></br>
     *  * TYPE_STRING = 0,
     *  * TYPE_FP = 1,
     *  * TYPE_INT = 2,
     *  * TYPE_FORMULA = 3,
     *  * TYPE_BOOLEAN = 4,
     *  * TYPE_DOUBLE = 5;
     *
     * @return int type for this Cell
     */
    override val cellType: Int
        get() = record!!.cellType

    /**
     * return the underlying cell record <br></br>
     * for internal API use only
     *
     * @return XLS cell record
     */
    val cell: BiffRec?
        get() = record

    /**
     * returns Border Colors of Cell in: top, left, bottom, right order <br></br>
     * returns null or a java.awt.Color object for each of 4 sides
     *
     * @return java.awt.Color array representing the 4 borders of the cell
     */
    val borderColors: Array<Color>
        get() {
            formatHandle
            return this.formatter!!.borderColors
        }

    /**
     * returns the low-level bytes for the underlying BIFF8 record. <br></br>
     * For Internal API use only
     *
     * @return bytes for the underlying record
     */
    val bytes: ByteArray?
        get() = record!!.data

    /**
     * Returns the column number of this Cell.
     *
     * @return int the Column Number of the Cell
     */
    override val colNum: Int
        get() {
            setMulblank()
            return record!!.colNumber.toInt()
        }

    /**
     * Returns the row number of this Cell.
     *
     *
     * NOTE: This is the 1-based row number such as you will see in a spreadsheet
     * UI.
     *
     *
     * ie: A1 = row 1
     *
     * @return 1-based int the Row Number of the Cell
     */
    override val rowNum: Int
        get() = record!!.rowNumber

    /**
     * Returns the Address of this Cell as a String.
     *
     * @return String the address of this Cell in the WorkSheet
     */
    override val cellAddress: String
        get() {
            setMulblank()
            return record!!.cellAddress

        }

    val intLocation: IntArray
        get() {
            setMulblank()
            return record!!.intLocation
        }

    /**
     * Returns the Address of this Cell as a String. Includes the sheet name in the
     * address.
     *
     * @return String the address of this Cell in the WorkSheet
     */
    val cellAddressWithSheet: String
        get() {
            setMulblank()
            return record!!.cellAddressWithSheet

        }

    /**
     * sets the column number referenced in the set of multiple blanks <br></br>
     * for multiple blank cells only <br></br>
     * for Internal Use. Not intended for the End User.
     */
    internal var mulblankcolnum: Short = -1

    /**
     * Returns the name of this Cell's WorkSheet as a String.
     *
     * @return String the name this Cell's WorkSheet
     */
    override val workSheetName: String
        get() {
            if (workSheetHandle == null) {
                try {
                    return record!!.sheet!!.sheetName
                } catch (e: Exception) {
                    return ""
                }

            }
            return workSheetHandle!!.sheetName
        }

    /**
     * Returns the value of the Cell as a String. <br></br>
     * boolean Cell types will return "true" or "false"
     *
     * @return String the value of the Cell
     */
    /**
     * set the value of this cell to String s <br></br>
     * NOTE: this method will not check for formula references or do any data
     * conversions <br></br>
     * This method is useful when a string may start with = but you do not want to
     * convert to a Formula value
     *
     *
     * If you wish to put a line break in the string use the newline "\n" character.
     * Note this will not function unless you also apply a format handle to the cell
     * with WrapText=true
     *
     * @param String s - the String value to set the Cell to
     * @throws CellTypeMismatchException
     * @see setVal
     */
    // avoid potential issues with string
    // values beginning with "="
    var stringVal: String?
        get() = record!!.stringVal
        set(s) = try {
            if ((s == null || s == "") && record !is Blank)
                changeCellType(s)
            else if (s != null && s != "") {
                if (record !is Labelsst)
                    changeCellType(" ")
                record!!.stringVal = s
            }
        } catch (e: Exception) {
            throw CellTypeMismatchException(e.toString())
        }

    /**
     * Gets the value of the cell as a String with the number format applied.
     * Boolean cell types will return "true" or "false". Custom number formats are
     * not currently supported, although they will be written correctly to the
     * output file. Patterns that display negative numbers in red are not currently
     * supported; the number will be prefixed with a minus sign instead. For more
     * information on number format patterns see
     * [Microsoft KB264372](http://support.microsoft.com/kb/264372).
     *
     * @return the value of the cell as a string formatted according to the cell
     * type and, if present, the number format pattern
     */
    override val formattedStringVal: String
        get() {
            val myfmt = this.formatHandle
            return CellFormatFactory.fromPatternString(myfmt!!.formatPattern).format(this)
        }

    /**
     * Returns the Formatting record ID (FormatId) for this Cell <br></br>
     * This can be used with 'setFormatId(int i)' to copy the formatting from one
     * Cell to another (e.g. a template cell to a new cell)
     *
     * @return int the FormatId for this Cell
     */
    /**
     * Sets the Formatting record ID (FormatId) for this Cell
     *
     *
     * This can be used with 'getFormatId()' to copy the formatting from one Cell to
     * another (ie: a template cell to a new cell)
     *
     * @param int i - the new index to the Format for this Cell
     */
    override var formatId: Int
        get() {
            setMulblank()
            return record!!.ixfe
        }
        set(i) = record!!.setXFRecord(i)

    /**
     * get the conditional formatting record ID (FormatId) <br></br>
     * returns the normal FormatId if the condition(s) in the conditional format
     * have not been met
     *
     * @return int the conditional format id
     */
    // TODO: only supporting first cfmat handle again
    // create a ptgref for this Cell
    // TODO: evaluate and combine multiple rules...
    // currently returns on first true format
    val conditionalFormatId: Int
        get() {
            val cfhandles = this.conditionalFormatHandles
            if (cfhandles == null || cfhandles.size == 0) {
                return this.formatId
            }
            val cfmt = cfhandles[0].cndfmt

            val x = cfmt!!.rules.iterator()
            while (x.hasNext()) {
                val cx1 = x.next() as Cf
                val pref = PtgRef(this.cellAddress, this.record, false)

                if (cx1.evaluate(pref)) {
                    return cfmt.cfxe
                }
            }
            return this.formatId
        }

    /**
     * returns an array of FormatHandles for the Cell that have the current
     * conditional format applied to them.
     *
     *
     * This behavior is still in testing, and may change
     *
     * @return an array of FormatHandles, one for each of the Conditional Formatting
     * rules
     */
    // TODO, should these be read-only?
    // TODO, this is bad, only handles first cf record for the cell
    val conditionallyFormattedHandles: Array<FormatHandle>?
        get() {
            val cfhandles = this.conditionalFormatHandles
            if (cfhandles != null) {
                val fmx = arrayOfNulls<FormatHandle>(cfhandles[0].cndfmt!!.rules.size)
                var t = 0
                while (t < fmx.size) {
                    fmx[t++] = FormatHandle(cfhandles[0].cndfmt!!, workBook!!, t, this)
                    t++
                }
                return fmx
            }
            return null
        }

    /**
     * return all the ConditionalFormatHandles for this Cell, if any
     *
     * @return
     */
    val conditionalFormatHandles: Array<ConditionalFormatHandle>?
        get() {
            val sh = this.workSheetHandle ?: return null
            val cfs = sh.conditionalFormatHandles
            val cfhandles = ArrayList<ConditionalFormatHandle>()
            for (i in cfs.indices) {
                if (cfs[i].contains(this))
                    cfhandles.add(cfs[i])
            }
            val c = arrayOfNulls<ConditionalFormatHandle>(cfhandles.size)
            return cfhandles.toTypedArray<ConditionalFormatHandle>()
        }

    /**
     * Gets the FormatHandle (a Format Object describing the formats for this Cell)
     * for this Cell.
     *
     * @return FormatHandle for this Cell
     */
    /**
     * Sets the FormatHandle (a Format Object describing the formats for this Cell)
     * for this Cell
     *
     * @param FormatHandle to apply to this Cell
     * @see FormatHandle
     */
    var formatHandle: FormatHandle?
        get() {
            if (this.formatter == null)
                this.setFormatHandle()

            return this.formatter
        }
        set(f) {
            f.addCell(this.record)
            this.formatter = f
        }

    /**
     * Returns the Formula Handle (an Object describing a Formula) for this Cell, if
     * it contains a formula
     *
     * @return FormulaHandle the Formula of the Cell
     * @throws FormulaNotFoundException
     * @see FormulaHandle
     */
    val formulaHandle: FormulaHandle
        @Throws(FormulaNotFoundException::class)
        get() {
            val f = record!!.formulaRec ?: throw FormulaNotFoundException("No Formula for: $cellAddress")
            return FormulaHandle(f, workBook)
        }

    /**
     * Returns the Hyperlink URL String for this Cell, if any
     *
     * @return String URL if this Cell contains a hyperlink
     */
    /**
     * Creates a new Hyperlink for this Cell from a URL String. Can be any valid
     * URL. This URL String must include the protocol. <br></br>
     * For Example, "http://www.extentech.com/" <br></br>
     * To remove a hyperlink, pass in null for the URL String
     *
     * @param String urlstr - the URL String for this Cell
     */
    // TODO: remove existing Hlink from stream
    // mycell.hyperlink.remove(true);
    var url: String?
        get() = if (record!!.hyperlink != null) record!!.hyperlink!!.url else null
        set(urlstr) {
            if (urlstr == null) {
                record!!.hyperlink = null
                return
            }
            setURL(urlstr, "", "")
        }

    /**
     * Returns the URL Description String for this Cell, if any
     *
     * @return String URL Description, if this Cell contains a hyperlink
     */
    val urlDescription: String
        get() = if (record!!.hyperlink != null) record!!.hyperlink!!.description else ""

    /**
     * returns the value of this Cell as a double, if possible, or NaN if Cell value
     * cannot be converted to double
     *
     * @return double value or NaN if the Cell value cannot be converted to a double
     */
    val doubleVal: Double
        get() = record!!.dblVal

    /**
     * returns the value of this Cell as a int, if possible, or NaN if Cell value
     * cannot be converted to int
     *
     * @return int value or NaN if the Cell value cannot be converted to an int
     */
    val intVal: Int
        get() = record!!.intVal

    /**
     * returns the value of this Cell as a float, if possible, or NaN if Cell value
     * cannot be converted to float
     *
     * @return float value or NaN if the Cell value cannot be converted to an float
     */
    val floatVal: Float
        get() = record!!.floatVal

    /**
     * returns the value of this Cell as a boolean <br></br>
     * If the Cell is not of type Boolean, returns false
     *
     * @return boolean value of cell
     */
    val booleanVal: Boolean
        get() = record!!.booleanVal

    /**
     * get the index of the WorkSheet containing this Cell in the list of sheets
     *
     * @return int the WorkSheetHandle index for this Cell
     */
    val sheetNum: Int
        get() = this.record!!.sheet!!.sheetNum

    /**
     * Determines if the cellHandle represents a completely blank/null cell, and can
     * thus be ignored for many operations.
     *
     *
     * Criteria for returning true is a cell type of BLANK, that has a default
     * format id (0), is not part of a merge range, does not contain a URL, and is
     * not a part of a validation
     *
     * @return true if cell is truly blank
     */
    val isDefaultCell: Boolean
        get() = (this.cellType == CellHandle.TYPE_BLANK
                && (this.formatId == 15 && !this.workBook!!.workBook.isExcel2007 || this.formatId == 0)
                && this.mergedCellRange == null && this.url == null
                && this.validationHandle == null)

    /**
     * Returns an XML representation of the cell and it's component data
     *
     * @return String of XML
     */
    val xml: String
        get() = getXML(null)

    /**
     * Returns an int representing the current horizontal alignment in this Cell.
     *
     * @return int representing horizontal alignment
     * @see FormatHandle.ALIGN* constants
     */
    /**
     * Set the horizontal alignment for this Cell
     *
     * @param int align - constant value representing the horizontal alignment.
     * @see FormatHandle.ALIGN* constants
     */
    // 0 is default alignment
    var horizontalAlignment: Int
        get() = if (this.record!!.xfRec != null) {
            this.record!!.xfRec!!.horizontalAlignment
        } else 0
        set(align) {
            setFormatHandle()
            formatter!!.horizontalAlignment = align
        }

    /**
     * Returns an int representing the current vertical alignment in this Cell.
     *
     * @return int representing vertical alignment
     * @see FormatHandle.ALIGN* constants
     */
    /**
     * Set the Vertical alignment for this Cell
     *
     * @param int align - constant value representing the vertical alignment.
     * @see FormatHandle.ALIGN* constants
     */
    // 1 is default alignment
    var verticalAlignment: Int
        get() = if (this.record!!.xfRec != null) {
            this.record!!.xfRec!!.verticalAlignment
        } else 1
        set(align) {
            setFormatHandle()
            formatter!!.verticalAlignment = align
        }

    /**
     * Get the cell wrapping behavior for this cell.
     *
     * @return true if cell text is wrapped
     */
    /**
     * Sets the cell wrapping behavior for this cell
     *
     * @param boolean wrapit - true if cell text should be wrapped (default is false)
     */
    // false is default alignment
    // when wrap text it automatically wraps if row height has
    // NOT been set yet
    // has row height been altered??
    /* ignore */ var wrapText: Boolean
        get() = if (this.record!!.xfRec != null) {
            this.record!!.xfRec!!.wrapText
        } else false
        set(wrapit) {
            setFormatHandle()
            formatter!!.wrapText = wrapit
            if (wrapit) {
                try {
                    if (!this.row.isAlteredHeight)
                        this.row.setRowHeightAutoFit()
                } catch (e: Exception) {
                }

            }
        }

    /**
     * Get the rotation of this Cell in degrees. <br></br>
     * Values 0-90 represent rotation up, 0-90degrees. <br></br>
     * Values 91-180 represent rotation down, 0-90 degrees. <br></br>
     * Value 255 is vertical
     *
     * @return int representing the degrees of cell rotation
     */
    /**
     * Set the rotation of the cell in degrees. <br></br>
     * Values 0-90 represent rotation up, 0-90degrees. <br></br>
     * Values 91-180 represent rotation down, 0-90 degrees. <br></br>
     * Value 255 is vertical
     *
     * @param int align - an int representing the rotation.
     */
    // false is default alignment
    var cellRotation: Int
        get() = if (this.record!!.xfRec != null) {
            this.record!!.xfRec!!.rotation
        } else 0
        set(align) {
            setFormatHandle()
            formatter!!.cellRotation = align
        }

    /**
     * retrieves or creates a new xf for this cell
     *
     * @return
     */
    private// reusing or creating new xfs is handled in FormatHandle/cloneXf and
    // updateXf
    // this.useExistingXF = true; // flag to re-use this XF
    // get the recidx of the last Xf
    // perform default add rec actions
    // update the pointer
    val newXf: Xf?
        get() {
            if (record!!.myxf != null)
                return record!!.myxf
            try {
                record!!.myxf = Xf(this.font!!.idx)
                val insertIdx = record!!.workBook!!.getXf(record!!.workBook!!.numXfs - 1)!!.recordIndex

                record!!.myxf!!.setSheet(null)
                record!!.workBook!!.streamer.addRecordAt(record!!.myxf!!, insertIdx + 1)
                record!!.workBook!!.addRecord(record!!.myxf!!, false)
                val xfe = record!!.myxf!!.idx
                record!!.ixfe = xfe
                return record!!.myxf
            } catch (e: Exception) {
                return null
            }

        }

    /**
     * Get the JSON object for this cell.
     *
     * @return String representing the JSON for this Cell
     */
    val json: String
        get() = jsonObject.toString()

    /**
     * Get the JSON object for this cell.
     */
    val jsonObject: JSONObject
        get() {
            val cr = mergedCellRange
            var mergedCellRange: IntArray? = null
            if (cr != null) {
                try {
                    mergedCellRange = cr.rangeCoords
                    if (record!!.opcode == XLSRecord.MULBLANK) {
                        val m = record as Mulblank?
                        if (!cr.contains(m!!.intLocation)) {
                            mergedCellRange = null
                        }
                    }
                } catch (e: CellNotFoundException) {
                }

            }

            return getJSONObject(mergedCellRange)
        }

    /**
     * Returns the validation handle for the cell.
     *
     * @return ValidationHandle for this Cell, or null if none
     */
    // somewhat normal?
    val validationHandle: ValidationHandle?
        get() {
            var ret: ValidationHandle? = null
            try {
                ret = this.workSheetHandle!!.getValidationHandle(this.cellAddress)
            } catch (e: Exception) {
            }

            return ret
        }

    /**
     * Public Constructor added for use in Bean-context ONLY.
     *
     *
     * Do NOT create CellHandles manually.
     */
    constructor(c: BiffRec) {
        record = c as XLSRecord
    }

    /**
     * if this cellhandle refers to a mulblank, ensure internal mulblank properties
     * are set to appropriate cell in the mulblank range
     */
    private fun setMulblank() {
        if (record!!.opcode == XLSConstants.MULBLANK) {
            if (mulblankcolnum.toInt() == -1) { // init
                mulblankcolnum = record!!.colNumber
                (record as Mulblank).setCurrentCell(mulblankcolnum)
                record!!.ixfe // ensure myxf is set to correct
                // xf for the given cell in the
                // set of mulblanks
            } else if (mulblankcolnum != record!!.colNumber) {
                (record as Mulblank).setCurrentCell(mulblankcolnum)
                record!!.ixfe // ensure myxf is set to correct
                // xf for the given cell in the
                // set of mulblanks
            }
            this.formatter = null
        }
    }

    /**
     * Public Constructor added for use in Bean-context ONLY.
     *
     *
     * Do NOT create CellHandles manually.
     */
    constructor(c: BiffRec, myb: WorkBook) {
        record = c as XLSRecord
        setMulblank()
        this.workBook = myb
    }

    /**
     * Get a FormatHandle (a Format Object describing the formats for this Cell)
     * referenced by this CellHandle.
     *
     * @return FormatHandle
     * @see FormatHandle
     */
    internal fun setFormatHandle() {
        setMulblank()
        if (formatter != null && formatter!!.formatId == this.record!!.ixfe) {
            return
        }
        // reusing or creating new xfs is handled in FormatHandle/cloneXf and
        // updateXf
        if (this.record!!.xfRec != null) {
            formatter = FormatHandle(this.workBook, this.record!!.myxf)
        } else {// should ever happen now?
            // useExistingXF = false;
            if (workBook == null && this.record!!.workBook != null)
                formatter = FormatHandle(this.record!!.workBook, -1)
            else
                formatter = FormatHandle(this.workBook!!, -1)
        }
        formatter!!.addCell(record)
    }

    /**
     * Sets a default "empty" value appropriate for the cell type of this CellHandle
     * <br></br>
     * For example, will set the value to 0.0 for TYPE_DOUBLE, an empty String for
     * TYPE_BLANK
     */
    fun setToDefault() {
        `val` = this.defaultVal
    }

    /**
     * Convenience method for toggling the bold state of the Font used by this Cell.
     *
     * @param boolean bold - true if bold
     */
    fun setBold(bold: Boolean) {
        setFormatHandle()
        formatter!!.bold = bold
    }

    /**
     * set the Foreground Color for this Cell. <br></br>
     * <br></br>
     * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
     * Background color is the PATTERN color for all patterns not equal to
     * PATTERN_SOLID <br></br>
     * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
     * Background Color is 64 (white).
     *
     * @param int t - Excel color constant
     * @see FormatHandle.COLOR constants
     */
    fun setForegroundColor(t: Int) {
        setFormatHandle()
        formatter!!.foregroundColor = t
    }

    /**
     * set the Foreground Color for this Cell <br></br>
     * NOTE: this is the PATTERN Color
     *
     * @param int Excel color constant
     * @see FormatHandle.COLOR constants
     */
    // TODO: is this doc correct?
    fun setForeColor(i: Int) {
        if (record!!.myxf == null)
            this.newXf
        record!!.myxf!!.setForeColor(i, null)
    }

    /**
     * set the Background Color for this Cell. <br></br>
     * <br></br>
     * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
     * Background color is the PATTERN color for all patterns not equal to
     * PATTERN_SOLID <br></br>
     * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
     * Background Color is 64 (white).
     *
     * @param int t - Excel color constant
     * @see FormatHandle.COLOR constants
     */
    fun setBackgroundColor(t: Int) {
        setFormatHandle()
        formatter!!.backgroundColor = t
    }

    /**
     * sets the Color of the Cell Foreground pattern for this Cell.
     *
     * <br></br>
     * <br></br>
     * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
     * Background color is the PATTERN color for all patterns not equal to
     * PATTERN_SOLID <br></br>
     * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
     * Background Color is 64 (white).
     *
     * @param java .awt.Color col - color for the foreground
     */
    fun setForegroundColor(col: Color) {
        setFormatHandle()
        formatter!!.setForegroundColor(col)
    }

    /**
     * set the Color of the Cell Background pattern for this Cell. <br></br>
     * <br></br>
     * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
     * Background color is the PATTERN color for all patterns not equal to
     * PATTERN_SOLID <br></br>
     * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
     * Background Color is 64 (white).
     *
     * @param java .awt.Color col - background color
     */
    fun setBackgroundColor(col: Color) {
        setFormatHandle()
        formatter!!.setBackgroundColor(col)
    }

    /**
     * set the Color of the Cell Background for this Cell. <br></br>
     * <br></br>
     * see FormatHandle.COLOR constants for valid values <br></br>
     * <br></br>
     * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
     * Background color is the PATTERN color for all patterns not equal to
     * PATTERN_SOLID <br></br>
     * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
     * Background Color is 64 (white).
     *
     * @param int t - Excel color constant for Cell Background color
     */
    fun setCellBackgroundColor(t: Int) {
        setFormatHandle()
        formatter!!.cellBackgroundColor = t
    }

    /**
     * set the Color of the Cell Background for this Cell. <br></br>
     * <br></br>
     * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
     * Background color is the PATTERN color for all patterns not equal to
     * PATTERN_SOLID <br></br>
     * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
     * Background Color is 64 (white).
     *
     * @param java .awt.Color col - color of the cell background
     */
    fun setCellBackgroundColor(col: Color) {
        setFormatHandle()
        formatter!!.setCellBackgroundColor(col)
    }

    /**
     * set the Color of the Cell Background Pattern for this Cell.
     *
     * @param java .awt.Color col - color of the pattern background
     */
    fun setPatternBackgroundColor(col: Color) {
        setFormatHandle()
        formatter!!.setBackgroundColor(col)
    }

    /**
     * set the Color of the Border for this Cell.
     *
     * @param java .awt.Color col - border color
     */
    fun setBorderColor(col: Color) {
        setFormatHandle()
        formatter!!.setBorderColor(col)
    }

    /**
     * set the Color of the right Border line for this Cell.
     *
     * @param java .awt.Color col - right border color
     */
    fun setBorderRightColor(col: Color) {
        setFormatHandle()
        formatter!!.borderRightColor = col
    }

    /**
     * set the Color of the left Border line for this Cell.
     *
     * @param java .awt.Color col - left border color
     */
    fun setBorderLeftColor(col: Color) {
        setFormatHandle()
        formatter!!.borderLeftColor = col
    }

    /**
     * set the Color of the top Border line for this Cell.
     *
     * @param java .awt.Color col - top border color
     */
    fun setBorderTopColor(col: Color) {
        setFormatHandle()
        formatter!!.borderTopColor = col
    }

    /**
     * set the Color of the bottom Border line for this Cell.
     *
     * @param java .awt.Color col - bottom border color
     */
    fun setBorderBottomColor(col: Color) {
        setFormatHandle()
        formatter!!.borderBottomColor = col
    }

    /**
     * set the Border line style for this Cell.
     *
     * @param short s - border constant
     * @see FormatHandle.BORDER line style constants
     */
    fun setBorderLineStyle(s: Short) {
        setFormatHandle()
        formatter!!.setBorderLineStyle(s)
    }

    /**
     * set the Right Border line style for this Cell.
     *
     * @param short s - border constant
     * @see FormatHandle.BORDER line style constants
     */
    fun setRightBorderLineStyle(s: Short) {
        setFormatHandle()
        formatter!!.rightBorderLineStyle = s
    }

    /**
     * set the Left Border line style for this Cell.
     *
     * @param short s - border constant
     * @see FormatHandle.BORDER line style constants
     */
    fun setLeftBorderLineStyle(s: Short) {
        setFormatHandle()
        formatter!!.leftBorderLineStyle = s
    }

    /**
     * set the Top Border line style for this Cell.
     *
     * @param short s - border constant
     * @see FormatHandle.BORDER line style constants
     */
    fun setTopBorderLineStyle(s: Short) {
        setFormatHandle()
        formatter!!.topBorderLineStyle = s
    }

    /**
     * set the Bottom Border line style for this Cell.
     *
     * @param short s - border constant
     * @see FormatHandle.BORDER line style constants
     */
    fun setBottomBorderLineStyle(s: Short) {
        setFormatHandle()
        formatter!!.bottomBorderLineStyle = s
    }

    /**
     * removes the borders for this cell
     */
    fun removeBorder() {
        setFormatHandle()
        formatter!!.removeBorders()
    }

    /**
     * Set the Font for this Cell via font name, font style and font size <br></br>
     * This method adds a new Font to the WorkBook. <br></br>
     * Roughly matches the functionality of the java.awt.Font class.
     *
     * @param String fn - system name of the font e.g. "Arial"
     * @param int    stl - font style (either Font.BOLD or Font.PLAIN)
     * @param int    sz - font size in points
     */
    fun setFont(fn: String, stl: Int, sz: Int) {
        setFormatHandle()
        formatter!!.setFont(fn, stl, sz.toDouble())
    }

    /**
     * Removes any note/comment records attached to this cell
     */
    fun removeComment() {
        try {
            val note = this.comment
            note.remove()
        } catch (e: DocumentObjectNotFoundException) {
        }

    }

    /**
     * Creates a new annotation (Note or Comment) to the cell
     *
     * @param comment -- text of note
     * @param author  -- name of author
     * @return CommentHandle - handle which allows access to the Note object
     * @see CommentHandle
     */
    fun createComment(comment: String, author: String): CommentHandle {
        val n = record!!.sheet!!.createNote(this.cellAddress, comment, author)
        return CommentHandle(n)
    }

    /**
     * set the Font Color for this Cell <br></br>
     * <br></br>
     * see FormatHandle.COLOR constants for valid values
     *
     * @param int t - Excel color constant
     */
    fun setFontColor(t: Int) {
        setFormatHandle()
        formatter!!.setFontColor(t)
    }

    fun setBlankRef(c: Int) {
        mulblankcolnum = c.toShort()
    }

    /**
     * Gets the value of the cell as a String with the number format applied.
     * Boolean cell types will return "true" or "false". Custom number formats are
     * not currently supported, although they will be written correctly to the
     * output file. Patterns that display negative numbers in red are not currently
     * supported; the number will be prefixed with a minus sign instead. For more
     * information on number format patterns see
     * [Microsoft KB264372](http://support.microsoft.com/kb/264372).
     *
     * @param formatForXML if true non-compliant characters will be properly qualified
     * @return the value of the cell as a string formatted according to the cell
     * type and, if present, the number format pattern
     */
    fun getFormattedStringVal(formatForXML: Boolean): String {
        val myfmt = this.formatHandle
        var `val` = this.`val`!!.toString()
        if (formatForXML)
            `val` = io.starter.formats.XLS.OOXMLAdapter.stripNonAscii(`val`).toString()
        return CellFormatFactory.fromPatternString(myfmt!!.formatPattern).format(`val`)
    }

    /**
     * Returns the value of the Cell as a String. <br></br>
     * For numeric cell values, including cells containing a formula which return a
     * numeric value, the notation to be used in representing the value as a String
     * must be specified. <br></br>
     * The Notation choices are:
     *  * CellHandle.NOTATION_STANDARD e.g. "8002974342"
     *  * CellHandle.NOTATION_SCIENTIFIC e.g. "8.002974342E9"
     *  * CellHandle.NOTATION_SCIENTIFIC_EXCEL e.g. "8.002974342+E9" <br></br>
     * <br></br>
     * For non-numeric values, the value of the cell as a string is returned <br></br>
     * boolean Cell types will return "true" or "false"
     *
     * @param int notation one of the CellHandle.NOTATION constants for numeric cell
     * types; ignored for other cell types
     * @param int notation - notation constant
     * @return String the value of the Cell
     */
    fun getStringVal(notation: Int): String? {
        val numval = record!!.stringVal
        val i = this.cellType
        return if (i == TYPE_FP || i == TYPE_INT || i == TYPE_FORMULA || i == TYPE_DOUBLE) {
            ExcelTools.formatNumericNotation(numval!!, notation)
        } else numval
    }

    /**
     * Returns the value of the Cell as a String with the specified encoding. <br></br>
     * boolean Cell types will return "true" or "false"
     *
     * @param String encoding
     * @return String the value of the Cell
     */
    fun getStringVal(encoding: String): String? {
        return record!!.getStringVal(encoding)
    }

    /**
     * move this cell to another row <br></br>
     * throws CellPositionConflictException if there is a cell in that position
     * already
     *
     * @param int newrow - new row number
     * @throws CellPositionConflictException
     */
    @Throws(CellPositionConflictException::class)
    fun moveToRow(newrow: Int) {
        var newaddr = ExcelTools.getAlphaVal(record!!.colNumber.toInt())
        newaddr += newrow.toString()
        this.moveTo(newaddr)
    }

    /**
     * move this cell to another row <br></br>
     * overwrite any cells in the destination
     *
     * @param int newrow - new row number
     */
    fun moveAndOverwriteToRow(newrow: Int) {
        var newaddr = ExcelTools.getAlphaVal(record!!.colNumber.toInt())
        newaddr += newrow.toString()
        this.moveAndOverwriteTo(newaddr)
    }

    /**
     * move this cell to another column <br></br>
     * throws CellPositionConflictException if there is a cell in that position
     * already
     *
     * @param String newcol - the new column in alpha format e.g. "A", "B" ...
     * @throws CellPositionConflictException
     */
    @Throws(CellPositionConflictException::class)
    fun moveToCol(newcol: String) {
        var newaddr = newcol
        newaddr += (record!!.rowNumber + 1).toString()
        this.moveTo(newaddr)
    }

    /**
     * Copy all formats from a source Cell to this Cell
     *
     * @param CellHandle source - source cell
     */
    fun copyFormat(source: CellHandle) {
        this.cell!!.copyFormat(source.cell)
    }

    /**
     * copy this Cell to another location.
     *
     * @param String newaddr - address for copy of this Cell in Excel-style e.g. "A1"
     * @return returns the newly copied CellHandle
     * @throws CellPositionConflictException if there is a cell in the new address already
     */
    @Throws(CellPositionConflictException::class)
    fun copyTo(newaddr: String): CellHandle {

        // check for existing
        val bs = this.record!!.sheet

        val rec = this.record
        val nucell = (rec as XLSRecord).clone() as XLSRecord
        val rc = ExcelTools.getRowColFromString(newaddr)
        nucell.rowNumber = rc[0]
        nucell.setCol(rc[1].toShort())
        nucell.setXFRecord(this.record!!.ixfe)
        bs!!.addRecord(nucell, rc)

        val ret = CellHandle(nucell, workBook)
        if (record!!.hyperlink != null) {
            // set the bounds of the mycell.hyperlink
            ret.url = this.url
        }
        ret.workSheetHandle = this.workSheetHandle
        // this.getWorkSheetHandle().cellhandles.put(this.getCellAddress(),
        // this);
        return ret
    }

    /**
     * Removes this Cell from the WorkSheet
     *
     * @param boolean nullme - true if this CellHandle should be nullified after removal
     */
    fun remove(nullme: Boolean) {
        record!!.sheet!!.removeCell(record!!)
        if (nullme) {
            try {
                this.finalize()
            } catch (t: Throwable) {
            }

        }
    }

    /**
     * move this cell to another location.
     *
     * @param String newaddr - the new address for Cell in Excel-style notation e.g.
     * "A1"
     * @throws CellPositionConflictException if there is a cell in the new address already
     */
    @Throws(CellPositionConflictException::class)
    fun moveTo(newaddr: String) {

        // check for existing
        val bs = record!!.sheet
        val oldhand = bs!!.getCell(newaddr)
        if (oldhand != null)
            throw CellPositionConflictException(newaddr)
        bs.moveCell(this.cellAddress, newaddr)

        if (record!!.hyperlink != null) {
            // set the bounds of the mycell.hyperlink
            // int[] bnds =
            // ExcelTools.getRowColFromString(this.getCellAddress());

            val bnds = ExcelTools.getRowColFromString(this.cellAddress)

            val hl = record!!.hyperlink
            hl!!.rowFirst = bnds[0]
            hl.rowLast = bnds[0]
            hl.colFirst = bnds[1]
            hl.colLast = bnds[1]
            hl.init()
        }
    }

    /**
     * move this cell to another location, overwriting any cells that are in the way
     *
     * @param String newaddr - the new address for Cell in Excel-style notation e.g.
     * "A1"
     */
    fun moveAndOverwriteTo(newaddr: String) {

        // check for existing
        val bs = record!!.sheet
        val oldhand = bs!!.getCell(newaddr)
        bs.moveCell(this.cellAddress, newaddr)

        if (record!!.hyperlink != null) {
            val bnds = ExcelTools.getRowColFromString(this.cellAddress)

            val hl = record!!.hyperlink
            hl!!.rowFirst = bnds[0]
            hl.rowLast = bnds[0]
            hl.colFirst = bnds[1]
            hl.colLast = bnds[1]
            hl.init()
        }
    }

    /**
     * Set a conditional format upon this cell.
     *
     *
     * Note that conditional format handles are bound to a specific worksheet,
     *
     * @param format A ConditionalFormatHandle in the same worksheet
     */
    fun addConditionalFormat(format: ConditionalFormatHandle) {
        format.addCell(this)
    }

    /**
     * Resets this cell's format to the default.
     */
    fun clearFormats() {
        this.formatId = this.workBook!!.workBook.defaultIxfe
    }

    /**
     * Resets this cells contents to blank.
     */
    fun clearContents() {
        this.`val` = null
    }

    /**
     * Resets this cell to the default, as if it had just been added.
     */
    fun clear() {
        this.clearFormats()
        this.clearContents()
    }

    /**
     * returns true if this Cell contains a hyperlink
     *
     * @return boolean true if this Cell contains a hyperlink
     */
    fun hasHyperlink(): Boolean {
        return record!!.hyperlink != null
    }

    /**
     * Creates a new Hyperlink for this Cell from a URL String, a descrpiton and
     * textMark text. <br></br>
     * <br></br>
     * The URL String Can be any valid URL. This URL String must include the
     * protocol. For Example, "http://www.extentech.com/" <br></br>
     * The textMark text is the porition of the URL that follows # <br></br>
     * <br></br>
     * NOTE: URL text and textMark text must not be null or ""
     *
     * @param String urlstr - the URL String
     * @param String desc - the description text
     * @param String textMark - the text that follows #
     */
    fun setURL(urlstr: String, desc: String, textMark: String) {
        if (record!!.hyperlink != null) {
            record!!.hyperlink!!.setURL(urlstr, desc, textMark)
        } else {
            // create new URL
            record!!.hyperlink = Hlink.prototype as Hlink?
            record!!.hyperlink!!.setURL(urlstr, desc, textMark)

            // why would we want to set the val during this operation?
            // if (!desc.equals("")) setVal(desc);

            // set the bounds of the mycell.hyperlink
            val bnds = ExcelTools.getRowColFromString(this.cellAddress)
            record!!.hyperlink!!.rowFirst = bnds[0]
            record!!.hyperlink!!.colFirst = bnds[1]
            record!!.hyperlink!!.rowLast = bnds[0]
            record!!.hyperlink!!.colLast = bnds[1]
        }
    }

    /**
     * Sets a hyperlink to a location within the current template <br></br>
     * The URL String should be prefixed with "file://" <br></br>
     *
     * @param String fileURLStr - the file URL String
     */
    // TODO: document this: NOTE: Excel File URL in actuality does not match
    // documentation
    fun setFileURL(fileURLStr: String) {
        setFileURL(fileURLStr, "", "")
    }

    /**
     * Sets a hyperlink to a location within the current template, and includes
     * additional optional information: description + textMark text <br></br>
     * <br></br>
     * The URL String should be prefixed with "file://" <br></br>
     * <br></br>
     * The textMark text is the porition of the URL that follows #
     *
     * @param String fileURLstr - the file URL String
     * @param String desc - the description text
     * @param String textMark - text that follows #
     */
    // TODO: this documentation is contradictory
    fun setFileURL(fileURLstr: String, desc: String, textMark: String) {
        if (record!!.hyperlink != null) {
            record!!.hyperlink!!.setFileURL(fileURLstr, desc, textMark)
        } else {
            record!!.hyperlink = Hlink.prototype as Hlink?

            record!!.hyperlink!!.setFileURL(fileURLstr, desc, textMark)
            if (desc != "")
                `val` = desc

            // set the bounds of the mycell.hyperlink
            val bnds = ExcelTools.getRowColFromString(this.cellAddress)
            record!!.hyperlink!!.rowFirst = bnds[0]
            record!!.hyperlink!!.colFirst = bnds[1]
            record!!.hyperlink!!.rowLast = bnds[0]
            record!!.hyperlink!!.colLast = bnds[1]
        }
    }

    /**
     * set the value of this cell to Unicodestring us <br></br>
     * NOTE: This method will not check for formula references or do any data
     * conversions <br></br>
     * Useful when strings may start with = but you do not want to convert to a
     * formula value
     *
     * @param Unicodestring us - Unicode String
     * @throws CellTypeMismatchException
     */
    fun setStringVal(us: Unicodestring?) {
        try {
            if ((us == null || us.toString() == "") && record !is Blank)
                changeCellType(null) // set to blank
            else if (us != null && us.toString() != "") {
                if (record !is Labelsst)
                    changeCellType(" ") // avoid potential issues with string
                // values beginning with "="
                (record as Labelsst).setStringVal(us)
            }
        } catch (e: Exception) {
            throw CellTypeMismatchException(e.toString())
        }

    }

    /**
     * this method will be fired as each record is parsed from an input Spreadsheet
     *
     *
     * Dec 15, 2010
     */
    fun fireParserEvent() {

    }

    /**
     * Returns a String representation of this CellHandle
     *
     * @see java.lang.Object.toString
     */
    override fun toString(): String {
        var ret = this.cellAddress + ":" + this.stringVal
        if (this.url != null)
            ret += this.url
        return ret
    }

    /**
     * Set the Value of the Cell to a double
     *
     * @param double d- double value to set this Cell to
     * @throws CellTypeMismatchException
     */
    @Throws(CellTypeMismatchException::class)
    fun setVal(d: Double) {
        this.`val` = d
    }

    /**
     * Set value of this Cell to a Float
     *
     * @param float f - float value to set this Cell to
     * @throws CellTypeMismatchException
     */
    @Throws(CellTypeMismatchException::class)
    fun setVal(f: Float) {
        this.`val` = f
    }

    /**
     * Sets the value of this Cell to a java.sql.Date. <br></br>
     * You must also specify a formatting pattern for the new date, or null for the
     * default date format ("m/d/yy h:mm".) <br></br>
     * <br></br>
     * valid date format patterns: <br></br>
     * "m/d/y" <br></br>
     * "d-mmm-yy" <br></br>
     * "d-mmm" <br></br>
     * "mmm-yy" <br></br>
     * "h:mm AM/PM" <br></br>
     * "h:mm:ss AM/PM" <br></br>
     * "h:mm" <br></br>
     * "h:mm:ss" <br></br>
     * "m/d/yy h:mm" <br></br>
     * "mm:ss" <br></br>
     * "[h]:mm:ss" <br></br>
     * "mm:ss.0"
     *
     * @param java   .sql.Date dt - the value of the new Cell
     * @param String fmt - date formatting pattern
     */
    fun setVal(dt: java.sql.Date, fmt: String?) {
        var fmt = fmt

        if (this.workBook!!.formulaCalculationMode != WorkBook.CALCULATE_EXPLICIT)
            this.clearAffectedCells() // blow out cache
        if (fmt == null)
            fmt = "m/d/yyyy"
        this.`val` = DateConverter.getXLSDateVal(dt)
        this.formatPattern = fmt
    }

    /**
     * Sets the value of this Cell to a boolean value
     *
     * @param boolean b - boolean value to set this Cell to
     * @throws CellTypeMismatchException
     */
    @Throws(CellTypeMismatchException::class)
    fun setVal(b: Boolean) {
        `val` = java.lang.Boolean.valueOf(b)
    }

    /**
     * Set the value of this Cell to an int value <br></br>
     * NOTE: setting a Boolean Cell type to a zero or a negative number will set the
     * Cell to 'false'; setting it to an int value 1 or greater will set it to true.
     *
     * @param int i - int value to set this Cell to
     * @throws CellTypeMismatchException
     */
    @Throws(CellTypeMismatchException::class)
    fun setVal(i: Int) {
        if (record!!.cellType == XLSConstants.TYPE_BOOLEAN) {
            if (i > 0)
                `val` = java.lang.Boolean.valueOf(true)
            else
                `val` = java.lang.Boolean.valueOf(false)
        } else {
            `val` = Integer.valueOf(i)
        }
    }

    /**
     * Set a cell to Excel-compatible formula passed in as String. <br></br>
     *
     * @param String formStr - the Formula String
     * @throws FunctionNotSupportedException if unable to parse string correctly
     */
    @Throws(FunctionNotSupportedException::class)
    fun setFormula(formStr: String) {
        val ixfe = this.record!!.ixfe
        this.remove(true)
        this.record = workSheetHandle!!.add(formStr, this.cellAddress)!!.record
        this.record!!.setXFRecord(ixfe)
    }

    /**
     * Set a cell to formula passed in as String. Sets the cachedValue as well, so
     * no calculating is necessary.
     *
     *
     * Parses the string to convert into a Excel formula. <br></br>
     * IMPORTANT NOTE: if cell is ALREADY a formula String this method will NOT
     * reset it
     *
     * @param String formulaStr - The excel-compatible formula string to pass in
     * @param Object value - The calculated value of the formula
     * @throws Exception if unable to parse string correctly
     */
    @Throws(Exception::class)
    fun setFormula(formStr: String, value: Any) {
        if (this.record !is Formula) {
            val ixfe = this.record!!.ixfe
            val cr = this.record!!.mergeRange
            val r = this.rowNum
            val c = this.colNum
            this.remove(true)
            this.record = workSheetHandle!!.add(formStr, r, c, ixfe)!!.record
            this.record!!.mergeRange = cr
        }
        val f = this.record as Formula?
        f!!.setCachedValue(value)
    }

    /**
     * sets the formula for this cellhandle using a stack of Ptgs appropriate for
     * formula records. This method also sets the cachedValue of the formula as
     * well, so no new calculating is necessary.
     *
     * @param Stack  newExp - Stack of Ptgs
     * @param Object value - calculated value of formula
     */
    fun setFormula(newExp: Stack<*>, value: Any) {
        if (this.record !is Formula) {
            val ixfe = this.record!!.ixfe
            val mccr = this.record!!.mergeRange
            val r = this.rowNum
            val c = this.colNum
            this.remove(true)
            this.record = workSheetHandle!!.add("=0", r, c, ixfe)!!.record // add the most
            // basic formula so
            // can modify below
            // ((:
            this.record!!.mergeRange = mccr
        }
        try {
            val f = this.record as Formula?
            f!!.expression = newExp
            f.setCachedValue(value)
        } catch (e: Exception) {
            // do what??
        }

    }

    /**
     * Returns the size of the merged cell area, if one exists.
     *
     * @param row    this parameter is ignored
     * @param column this parameter is ignored
     * @return a 2 position int array with number of rows and number of cols
     */
    @Deprecated("since October 2012. This method duplicates the functionality of\n" +
            "      {@link #getMergedCellRange()}, which it calls internally. That\n" +
            "      method should be used instead.")
    fun getSpan(row: Int, column: Int): IntArray? {
        val mergerange = mergedCellRange
        if (mergerange != null) {
            if (DEBUG)
                Logger.logInfo("CellHandle $this getSpan() for range: $mergerange")
            val ret = intArrayOf(0, 0)
            // if(check.toString().equals(this.toString())){ //it's the first in
            // the range -- show it!
            try {
                ret[0] = mergerange.rows.size
                ret[1] = mergerange.cols.size // TODO: test!
            } catch (e: Exception) {
                Logger.logWarn("CellHandle getting CellSpan failed: $e")
            }

            // }
            return ret
        }
        return null
    }

    /**
     * Returns an XML representation of the cell and it's component data
     *
     * @param int[] mergedRange - include merged ranges in the XML representation if
     * not null
     * @return String of XML
     */
    fun getXML(mergedRange: IntArray?): String {
        val vl = ""
        var fvl = ""
        var sv = ""
        var hd = ""
        var csp = ""
        var hlink = ""
        var `val`: Any? = null
        val retval = StringBuffer()
        var typename = this.cellTypeName
        // put the formula string in
        if (typename == "Formula") {
            try {
                val fmh = formulaHandle
                val fms = fmh.formulaString
                // use single quotes around formula value to avoid errors in
                // xslt transform
                if (fms.indexOf("\"") > 0) {
                    fvl = " Formula='" + StringTool.convertXMLChars(fms) + "'"
                } else {
                    fvl = " Formula=\"" + StringTool.convertXMLChars(fms) + "\""
                }
                try {
                    if (this.workBook!!.workBook.calcMode != WorkBook.CALCULATE_EXPLICIT) {
                        `val` = fmh.calculate()
                    } else {
                        try {
                            // changed from getVal() now that getVal returns a
                            // null
                            `val` = fmh.stringVal
                        } catch (e: Exception) {
                            Logger.logWarn("CellHandle.getXML formula calc failed: $e")
                        }

                    }
                    if (`val` is Float)
                        typename = "Float"
                    else if (`val` is Double)
                        typename = "Double"
                    else if (`val` is Int)
                        typename = "Integer"
                    else if (`val` is java.util.Date || `val` is java.sql.Date
                            || `val` is java.sql.Timestamp)
                        typename = "DateTime"
                    else
                        typename = "String"
                } catch (e: Exception) {
                    typename = "String" // default
                }

            } catch (e: Exception) {
                Logger.logErr("OpenXLS.getXML() failed getting type of Formula for: $this", e)
            }

        }
        if (this.isDate)
            typename = "DateTime" // 20060428 KSC: Moved after Formula parsing

        // TODO: when RowHandle.getCells actually contains ALL cells, keep this
        if (this.record!!.opcode != XLSConstants.MULBLANK) {
            // Put the style ID in
            sv = " StyleID=\"s$formatId\""
            if (mergedRange != null) { // TODO: fix!
                csp += " MergeAcross=\"" + (mergedRange[3] - mergedRange[1] + 1) + "\""
                csp += " MergeDown=\"" + (mergedRange[2] - mergedRange[0]) + "\""
            }
            if (this.col.isHidden) {
                hd = " Hidden=\"true\""
            }
            // TODO: HRefScreenTip ????
            if (this.url != null) {
                hlink = " HRef=\"" + StringTool.convertXMLChars(this.url) + "\""
            }

            // put the date formattingin
            // Assemble the string
            retval.append("<Cell Address=\"" + this.cellAddress + "\"" + sv + csp + fvl + hd + hlink
                    + "><Data Type=\"" + typename + "\">")
            if (typename == "DateTime") {
                retval.append(DateConverter.getFormattedDateVal(this))
            } else if (this.cellType == CellHandle.TYPE_STRING) {
                `val` = this.stringVal // (String)getVal();
                if (`val` == "") { // 20070216 KSC: John, had the same edits,
                    // seems to work well in cursory tests ...
                    // retval.append(" "); does this screw up formulas expecting
                    // empty strings? -jm
                } else {
                    retval.append(StringTool.convertXMLChars(`val`!!.toString()))
                }
            } else {
                try {
                    // if(val == null)
                    `val` = this.`val`
                    retval.append(StringTool.convertXMLChars(`val`!!.toString()) + vl)
                } catch (e: Exception) {
                    Logger.logErr("CellHandle.getXML failed for: " + this.cellAddress + " in: "
                            + this.workBook!!.toString(), e)
                    retval.append("XML ERROR!")
                }

            }
            retval.append("</Data>")
            retval.append(end_cell_xml)
        } else {
            var c = (this.record as Mulblank).colFirst
            val lastcol = (this.record as Mulblank).colLast
            while (c <= lastcol) {
                mulblankcolnum = c.toShort()
                // Put the style ID in
                sv = " StyleID=\"s$formatId\""
                if (this.col.isHidden) {
                    hd = " Hidden=\"true\""
                }
                // TODO: HRefScreenTip ????
                if (this.url != null) {
                    hlink = " HRef=\"" + StringTool.convertXMLChars(this.url) + "\""
                }

                // put the date formattingin
                // Assemble the string
                retval.append("<Cell Address=\"" + this.cellAddress + "\"" + sv + csp + fvl + hd + hlink
                        + "><Data Type=\"" + typename + "\"/>")
                retval.append(end_cell_xml)
                c++
            }
        }
        return retval.toString()
    }

    override fun compareTo(that: CellHandle): Int {
        val comp = this.rowNum - that.rowNum
        return if (comp != 0) comp else this.colNum - that.colNum
    }

    override fun equals(that: Any?): Boolean {
        return if (that !is CellHandle) false else this.record == that.record
    }

    override fun hashCode(): Int {
        return this.record!!.hashCode()
    }

    /**
     * Set the super/sub script for the Font
     *
     * @param int ss - super/sub script constant (0 = none, 1 = super, 2 = sub)
     */
    fun setScript(ss: Int) {
        if (record!!.myxf == null)
            this.newXf
        record!!.myxf!!.font!!.script = ss
    }

    /**
     * Set the val of the biffrec with an Object
     *
     * @param Object to set the value of the Cell to
     */
    @Throws(CellTypeMismatchException::class)
    private fun setBiffRecValue(obj: Any?) {
        if (record!!.opcode == XLSConstants.BLANK || record!!.opcode == XLSConstants.MULBLANK) {
            // no reason for this Blank blank = (Blank)mycell;
            // String addr = mycell.getCellAddress();

            // trim the Mulblank
            /*
             * KSC: mulblanks are NOT expanded now if (blank.getMyMul() != null){ Mulblank
             * mblank = (Mulblank)blank.getMyMul(); mblank.trim(blank); }
             */
            changeCellType(obj) // 20080206 KSC: Basically replaces all above
            // code
        } else {
            if (obj == null) {
                // should never be false ??? if (!(mycell instanceof Blank))
                changeCellType(obj) // will set to blank
            } else if (obj is Float || obj is Double || obj is Int || obj is Long) {
                if (record is NumberRec || record is Rk) {
                    if (obj is Float) {
                        val f = obj as Float?
                        record!!.floatVal = f!!.toFloat()
                    } else if (obj is Int) {
                        val i = obj as Int?
                        record!!.intVal = i!!.toInt()
                    } else if (obj is Double) {
                        val d = obj as Double?
                        record!!.setDoubleVal(d!!.toDouble())
                    } else if (obj is Long) {
                        val d = obj as Long?
                        record!!.setDoubleVal(d!!.toLong().toDouble())
                    }
                } else
                    changeCellType(obj)
            } else if (obj is Boolean) {
                if (record is Boolerr)
                    record!!.booleanVal = obj.booleanValue()
                else
                    changeCellType(obj)
            } else if (obj is String) {
                if (obj.startsWith("="))
                    changeCellType(obj) // easier to just redo a formula...
                else if (!obj.toString().equals("", ignoreCase = true)) {
                    if (record is Labelsst)
                        record!!.stringVal = obj.toString()
                    else
                        changeCellType(obj)
                } else if (record !is Blank)
                    changeCellType(obj)
            }
        }
    }

    /**
     * if object type doesn't match current mycell record, remove and add
     * appropriate record type
     *
     * @param obj
     */
    private fun changeCellType(obj: Any?) {
        val rc = intArrayOf(record!!.rowNumber, record!!.colNumber.toInt())
        val bs = record!!.sheet
        val oldXf = record!!.ixfe
        bs!!.removeCell(record!!)
        val addedrec = bs.addValue(obj, rc, true)
        record = addedrec as XLSRecord
        record!!.setXFRecord(oldXf)
    }

    /**
     * Calculates and returns all formulas that reference this CellHandle. <br></br>
     * Please note that these cells may have already been calculated, so in order to
     * get their values without re-calculating them Extentech suggests setting the
     * book level non-calculation flag, ie
     * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT); or
     * FormulaHandle.getCachedVal()
     *
     * @return List of of calculated cells (CellHandles)
     */
    fun calculateAffectedCells(): List<CellHandle> {
        val rt = this.workBook!!.workBook.refTracker
        val its = rt!!.clearAffectedFormulaCells(this).values.iterator()

        val ret = ArrayList<CellHandle>()
        while (its.hasNext()) {
            val cx = CellHandle(its.next() as BiffRec, this.workBook)
            ret.add(cx)
        }
        return ret
    }

    /**
     * Internal method for clearing affected cells, does the same thing as
     * calculateAffectedCells, but does not create a list
     */
    fun clearAffectedCells() {
        val rt = this.workBook!!.workBook.refTracker
        rt!!.clearAffectedFormulaCells(this)
    }

    /**
     * Calculates and returns all formulas on the same sheet that reference this
     * CellHandle. <br></br>
     * Please note that these cells may have already been calculated, so in order to
     * get their values without re-calculating them Extentech suggests setting the
     * book level non-calculation flag, ie
     * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT); or
     * FormulaHandle.getCachedVal()
     *
     * @return List of of calculated cells (CellHandles)
     */
    fun calculateAffectedCellsOnSheet(): List<CellHandle> {
        val its = this.workBook!!.workBook.refTracker!!
                .clearAffectedFormulaCellsOnSheet(this, this.workSheetName).values.iterator()
        val ret = ArrayList<CellHandle>()
        while (its.hasNext()) {
            val cx = CellHandle(its.next() as BiffRec, this.workBook)
            ret.add(cx)
        }
        return ret
    }

    /**
     * flags chart references to the particular cell as dirty/ needing caches
     * rebuilt
     */
    fun clearChartReferences() {
        val ret = ArrayList<ChartHandle>()
        val ii = this.workBook!!.workBook.refTracker!!.getChartReferences(this.cell!!).iterator()
        while (ii.hasNext()) {
            val ai = ii.next() as Ai
            if (ai.parentChart != null)
                ai.parentChart!!.setMetricsDirty()
        }
    }

    /**
     * Get a JSON Object representation of a cell utilizing a merged range
     * identifier.
     *
     */
    @Deprecated("The {@code mergedRange} parameter is unnecessary. This method\n" +
            "      will be removed in a future version. Use {@link #getJSONObject()}\n" +
            "      instead.")
    fun getJSONObject(mergedRange: IntArray?): JSONObject {
        val theCell = JSONObject()
        try {
            theCell.put(JSON_LOCATION, cellAddress)

            var `val`: Any?
            try {
                `val` = `val`
                if (`val` == null)
                    `val` = ""
            } catch (ex: Exception) {
                Logger.logWarn("OpenXLS.getJSONObject failed: $ex")
                `val` = "#ERR!"
            }

            var typename = cellTypeName
            val dataval = JSONObject()

            if (typename == "Formula") {
                try {
                    val fmh = formulaHandle
                    val fms = fmh.formulaString

                    theCell.put(JSON_CELL_FORMULA, fms)

                    try {
                        if (java.lang.Float.parseFloat(`val`!!.toString()) == java.lang.Float.NaN) {
                            typename = JSON_FLOAT
                        } else if (`val` is Float)
                            typename = JSON_FLOAT
                        else if (`val` is Double)
                            typename = JSON_DOUBLE
                        else if (`val` is Int)
                            typename = JSON_INTEGER
                        else if (`val` is java.util.Date || `val` is java.sql.Date
                                || `val` is java.sql.Timestamp) {
                            typename = JSON_DATETIME
                        } else
                            typename = JSON_STRING
                    } catch (e: Exception) {
                        typename = JSON_STRING // default
                    }

                } catch (e: Exception) {
                    Logger.logErr("OpenXLS.getJSON() failed getting type of Formula for: " + toString(), e)
                }

            }

            if (isDate)
                typename = JSON_DATETIME

            dataval.put(JSON_TYPE, typename)

            // TODO: Handle Conditional Format
            // cell should return the style id for its condition
            // this is an ID that begins incrementing after the last Xf
            // and should be contained in the CSS for the output

            // if the conditional format evaluates to TRUE
            // then we use *that* style ID not the default

            // We can have multiple CF styles per cell, one per each rule... we'll need that
            // from CSS standpoint so...

            // style
            theCell.put(JSON_STYLEID, conditionalFormatId)

            // merges
            if (mergedRange != null) {
                theCell.put(JSON_MERGEACROSS, mergedRange[3] - mergedRange[1])
                theCell.put(JSON_MERGEDOWN, mergedRange[2] - mergedRange[0])
                if (isMergeParent) {
                    theCell.put(JSON_MERGEPARENT, true)
                } else {
                    theCell.put(JSON_MERGECHILD, true)
                }
            }

            // handle hidden setting
            try {
                if (col.isHidden)
                    theCell.put(JSON_HIDDEN, true)
            } catch (e: Exception) {
            }

            // handle the locked/formula hidden setting
            // only active if sheet is protected
            try {
                if (isFormulaHidden)
                    theCell.put(JSON_FORMULA_HIDDEN, true)

                theCell.put(JSON_LOCKED, isLocked)
            } catch (e: Exception) {
            }

            try {
                val vh = validationHandle
                if (vh != null)
                    theCell.put(JSON_VALIDATION_MESSAGE, vh.promptBoxTitle + ":" + vh.promptBoxText)
            } catch (e: Exception) {
            }

            // hyperlinks
            if (url != null)
                theCell.put(JSON_HREF, url)

            if (wrapText)
                theCell.put(JSON_WORD_WRAP, true)

            // store alignment for container issues
            val alignment = formatHandle!!.horizontalAlignment
            if (alignment == FormatHandle.ALIGN_RIGHT) {
                theCell.put(JSON_TEXT_ALIGN, "right")
            } else if (alignment == FormatHandle.ALIGN_CENTER) {
                theCell.put(JSON_TEXT_ALIGN, "center")
            } else if (alignment == FormatHandle.ALIGN_LEFT) {
                theCell.put(JSON_TEXT_ALIGN, "left")
            }

            // dates
            if (typename == JSON_DATETIME && `val` != null && `val` != "") {
                dataval.put(JSON_CELL_VALUE, formattedStringVal)
                // dataval.put(JSON_DATEVALUE, ch.getFloatVal());
                dataval.put("time", DateConverter.getCalendarFromCell(this).timeInMillis)
            } else if (cellType == CellHandle.TYPE_STRING) {
                // FORCES CALC
                if ((`val` as String).indexOf("\n") > -1) {
                    `val` = `val`.replace("\n".toRegex(), "<br/>")
                }
                if (`val` != "")
                    dataval.put(JSON_CELL_VALUE, `val`.toString())
            } else { // other
                dataval.put(JSON_CELL_VALUE, `val`!!.toString())
                try { // formatted pattern
                    val s = formatPattern
                    if (s != "") {
                        var fmtd = formattedStringVal // TRIGGERS CALC!
                        if (`val` != fmtd)
                            dataval.put(JSON_CELL_FORMATTED_VALUE, fmtd)
                        if (s!!.indexOf("Red") > -1) {
                            val d = Double(`val`.toString())
                            if (d < 0) {
                                theCell.put(JSON_RED_FORMAT, "1")
                                if (fmtd.indexOf("-") == 0)
                                    fmtd = fmtd.substring(1)
                                dataval.put(JSON_CELL_FORMATTED_VALUE, fmtd)
                            }
                        }
                    }
                } catch (x: Exception) {
                }

            }
            theCell.put(JSON_DATA, dataval)
        } catch (e: JSONException) {
            Logger.logErr("error getting JSON for the cell: $e")
        }

        return theCell
    }

    companion object {

        /**
         *
         */
        private const val serialVersionUID = 4737120893891570607L
        /**
         * Cell types
         */
        val TYPE_BLANK = Cell.TYPE_BLANK
        val TYPE_STRING = Cell.TYPE_STRING
        val TYPE_FP = Cell.TYPE_FP
        val TYPE_INT = Cell.TYPE_INT
        val TYPE_FORMULA = Cell.TYPE_FORMULA
        val TYPE_BOOLEAN = Cell.TYPE_BOOLEAN
        val TYPE_DOUBLE = Cell.TYPE_DOUBLE
        val NOTATION_STANDARD = 0
        val NOTATION_SCIENTIFIC = 1
        val NOTATION_SCIENTIFIC_EXCEL = 2

        internal val begin_hidden_emptycell_xml = "<Cell Address=\""
        internal val end_hidden_emptycell_xml = "\" StyleID=\"s15\" Hidden=\"true\"><Data Type=\"String\"></Data></Cell>"

        internal val begin_cell_xml = "<Cell Address=\""
        internal val end_emptycell_xml = "\" StyleID=\"s15\"><Data Type=\"String\"></Data></Cell>"
        internal val end_cell_xml = "</Cell>"

        /**
         * Returns an xml representation of an empty cell
         *
         * @param loc       - the cell address
         * @param isVisible - if the cell is visible (not hidden)
         * @return
         */
        protected fun getEmptyCellXML(loc: String, isVisible: Boolean): String {
            return if (!isVisible) {
                begin_hidden_emptycell_xml + loc + end_hidden_emptycell_xml
            } else {
                begin_cell_xml + loc + end_emptycell_xml
            }
        }

        /**
         * Creates a copy of this cell on the given worksheet at the given address.
         *
         * @param sourcecell  the cell to copy
         * @param newsheet    the sheet to which the cell should be copied
         * @param row         the row in which the copied cell should be placed
         * @param col         the row in which the copied cell should be placed
         * @param copyByValue whether to copy formulas' values instead of the formulas
         * themselves
         * @return CellHandle representing the new cell
         */
        @JvmOverloads
        fun copyCellToWorkSheet(sourcecell: CellHandle, newsheet: WorkSheetHandle, row: Int, col: Int,
                                copyByValue: Boolean = false): CellHandle {
            // copy cell values
            var newcell: CellHandle? = null

            val offsets = intArrayOf(row - sourcecell.rowNum, col - sourcecell.colNum)

            if (sourcecell.isFormula && !copyByValue)
                try {
                    val fmh = sourcecell.formulaHandle
                    newcell = newsheet.add(fmh.formulaString, row, col)
                    val fm2 = newcell!!.formulaHandle
                    FormulaHandle.moveCellRefs(fm2, offsets)
                } catch (ex: FormulaNotFoundException) {
                    newcell = null
                }

            if (newcell == null)
                newcell = newsheet.add(sourcecell.`val`, row, col)

            return copyCellHelper(sourcecell, newcell!!)
        }

        /**
         * Create a copy of this Cell in another WorkBook
         *
         * @param sourcecell the cell to copy
         * @param target     worksheet to copy this cell into
         * @return
         */
        fun copyCellToWorkSheet(sourcecell: CellHandle, newsheet: WorkSheetHandle): CellHandle {
            // copy cell values
            var newcell: CellHandle? = null
            try {
                val fmh = sourcecell.formulaHandle
                // Logger.logInfo("testFormats Formula encountered: "+
                // fmh.getFormulaString());

                newcell = newsheet.add(fmh.formulaString, sourcecell.cellAddress)
            } catch (ex: FormulaNotFoundException) {
                newcell = newsheet.add(sourcecell.`val`, sourcecell.cellAddress)
            }

            return copyCellHelper(sourcecell, newcell!!)
        }

        /**
         * Copies all formatting - xf and non-xf (such as column width, hidden state)
         * plus merged cell range from a sourcecell to a new cell (usually in a new
         * workbook)
         *
         * @param sourcecell the cell to copy
         * @param newcell    the cell to copy to
         * @return
         */
        protected fun copyCellHelper(sourcecell: CellHandle, newcell: CellHandle): CellHandle {
            // copy row height & attributes
            val rz = sourcecell.row.height
            newcell.row.height = rz
            if (sourcecell.row.isHidden) {
                newcell.row.isHidden = true
            }
            // copy col width & attributes
            val rzx = sourcecell.col.width
            newcell.col.width = rzx
            if (sourcecell.col.isHidden) {
                newcell.col.isHidden = true
                // Logger.logInfo("column " + rzx + " is hidden");
            }

            try {
                // copy merged ranges
                var rng = sourcecell.mergedCellRange
                if (rng != null) {
                    rng = CellRange(rng.getRange(), newcell.workBook)
                    rng.addCellToRange(newcell)
                    rng.mergeCells(false)
                }
                // Handle formats:
                val origxf = sourcecell.workBook!!.workBook.getXf(sourcecell.formatId)
                newcell.formatHandle!!.addXf(origxf!!)
                return newcell
            } catch (ex: Exception) {
                Logger.logErr("CellHandle.copyCellHelper failed.", ex)
            }

            return newcell
        }
    }

}
/**
 * Creates a copy of this cell on the given worksheet at the given address.
 *
 * @param sourcecell the cell to copy
 * @param newsheet   the sheet to which the cell should be copied
 * @param row        the row in which the copied cell should be placed
 * @param col        the row in which the copied cell should be placed
 * @return CellHandle representing the new Cell
 */
