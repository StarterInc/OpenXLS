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

import io.starter.formats.XLS.*
import io.starter.toolkit.StringTool

import java.util.ArrayList


/**
 * The ColHandle provides access to an Worksheet Column and its Cells.
 * <br></br>
 * Use the ColHandle to work with individual Columns in an XLS file.
 * <br></br>
 * With a ColHandle you can:
 * <br></br><blockquote>
 * get a handle to the Cells in a column<br></br>
 * set the default formatting for a column<br></br>
 * <br></br></blockquote>
 *
 *
 * Note: for a discussion of Column widths see:
 * http://support.microsoft.com/?kbid=214123
 *
 * @see WorkBookHandle
 *
 * @see WorkSheetHandle
 *
 * @see FormulaHandle
 */
class ColHandle
/**
 * creates a new  ColHandle from a Colinfo Object and reference to a worksheet (WorkSheetHandle Object)
 *
 * @param c
 * @param sheet
 */
(private val myCol: Colinfo, private val mySheet: WorkSheetHandle) {
    private var formatter: FormatHandle? = null
    private val wbh: WorkBook?

    private var lastsz = 0 // the last checked col width

    /**
     * returns the width of this Column in internal units
     * defined as follows:
     * <br></br>
     * default width of the columns in 1/256 of the width of the zero character,
     * using default font.
     * <br></br>The Default Excel Column, whose width in Excel Units or Characters is 8.43, has a width in these units of 2300.
     *
     * @return int Column width in internal units
     */
    /**
     * sets the width of this Column in internal units, described as follows:
     * <br></br>
     * default width of the columns in 1/256 of the width of the zero character,
     * using default font.
     * <br></br>The Default Excel Column, whose width in Characters or Excel Units, is 8.43, has a width in these units of 2300.
     *
     * NOTE:
     * The last Cell in the column having its width
     * set will be the resulting width of the column
     *
     * @param int i - desired Column width in internal units
     */
    /* if an image falls upon this column,
         * adjust image width so that it does not change
         */// for each image that falls over this column, trap index + original width -- to be reset after setting col width
    // should only be one, right?
    // now adjust any of the images that we noted above
    var width: Int
        get() = myCol.colWidth
        set(newWidth) {
            val iAdjust = ArrayList<IntArray>()
            val images = myCol.sheet!!.images
            if (images != null) {
                for (z in images.indices) {
                    val ih = images[z]
                    val c0 = ih.col
                    val c1 = ih.col1
                    val col = myCol.colFirst
                    if (col >= c0 && col <= c1) {
                        val w = ih.width.toInt()
                        iAdjust.add(intArrayOf(z, w))
                    }
                }
            }
            lastsz = newWidth
            myCol.colWidth = newWidth
            for (z in iAdjust.indices) {
                val ih = images!![iAdjust[z][0]]
                ih.setWidth(iAdjust[z][1])
            }
        }


    /**
     * returns the width of this Column in Characters or regular Excel units
     * <br></br>NOTE: this value is a calculated value that should be close but still is an approximation of Excel units
     *
     * @return int Column width in Excel units
     */
    /**
     * sets the width of this Column in Characters or Excel units.
     * <br></br>
     * The default Excel column width is set to 8.43 Characters,
     * based on the default font and font size,
     * <br></br>
     * NOTE: The last Cell in the column having its width
     * set will be the resulting width of the column
     *
     * @param int i - desired Column width in Characters (Excel units)
     */
    /* if an image falls upon this column,
         * adjust image width so that it does not change
         */// for each image that falls over this column, trap index + original width -- to be reset after setting col width
    // should only be one, right?
    var widthInChars: Int
        get() = myCol.colWidthInChars
        set(newWidth) {
            val iAdjust = ArrayList<IntArray>()
            val images = myCol.sheet!!.images
            if (images != null) {
                for (z in images.indices) {
                    val ih = images[z]
                    val c0 = ih.col
                    val c1 = ih.col1
                    val col = myCol.colFirst
                    if (col >= c0 && col <= c1) {
                        val w = ih.width.toInt()
                        iAdjust.add(intArrayOf(z, w))
                    }
                }
            }

            myCol.setColWidthInChars(newWidth.toDouble())
            for (z in iAdjust.indices) {
                val ih = images!![iAdjust[z][0]]
                ih.setWidth(iAdjust[z][1])
            }
        }

    /**
     * returns the format ID (the index to the format record) for this Column
     * <br></br>The Column format is the default formatting for each cell contained
     * within the column
     *
     * @return int formatId - the index of the format record for this Column
     * @see FormatHandle
     */
    /**
     * sets the format id (an index to a Format record) for this Column
     * <br></br>This sets the default formatting for the Column
     * such that any cell that does not specifically set it's own formatting
     * will display this Column formatting
     *
     * @param int i - ID representing the Format to set this Column
     * @see FormatHandle
     */
    var formatId: Int
        get() = myCol.ixfe
        set(i) {
            myCol.ixfe = i
        }

    /**
     * returns the FormatHandle (a Format Object describing visual properties) for this Column
     * <br></br>NOTE: The Column format record describes the default formatting for each cell contained
     * within the column
     *
     * @return FormatHandle - a Format object to apply to this Col
     */
    val formatHandle: FormatHandle?
        get() {
            if (this.formatter == null) this.setFormatHandle()
            return this.formatter
        }

    /**
     * returns the first Column referenced by this column handle
     * <br></br>NOTE: A Column handle may in some circumstances refer to a range of columns
     *
     * @return int first column number referenced by this Column handle
     */
    val colFirst: Int
        get() = myCol.colFirst

    /**
     * returns the last Column referenced by this column handle
     * <br></br>NOTE: A Column handle may in some circumstances refer to a range of columns
     *
     * @return int last column number referenced by this Column handle
     */
    val colLast: Int
        get() = myCol.colLast


    /**
     * returns the array of Cells in this Column
     *
     * @return CellHandle array
     */
    val cells: Array<CellHandle>
        get() {
            val mycells: List<*>
            try {
                mycells = this.mySheet.boundsheet!!.getCellsByCol(this.colFirst)
            } catch (e: CellNotFoundException) {
                return arrayOfNulls(0)
            }

            val ch = arrayOfNulls<CellHandle>(mycells.size)
            for (t in ch.indices) {
                ch[t] = CellHandle(mycells[t] as BiffRec, null)
                ch[t].workSheetHandle = null
            }
            return ch
        }

    /**
     * Returns the Outline level (depth) of this Column
     *
     * @return int outline level
     */
    /**
     * Set the Outline level (depth) of this Column
     *
     * @param int x - outline level
     */
    var outlineLevel: Int
        get() = myCol.outlineLevel
        set(x) {
            this.myCol.outlineLevel = x
        }

    /**
     * returns true if this Column is collapsed
     *
     * @return true if ths Column is collapsed, false otherwise
     */
    /**
     * sets whether to collapse this Column
     *
     * @param boolean b - true to collapse this Column
     */
    var isCollapsed: Boolean
        get() = myCol.isCollapsed
        set(b) {
            this.myCol.isCollapsed = b
        }

    /**
     * returns true if this Column is hidden
     *
     * @return true if this Column is hidden, false if not
     */
    /**
     * sets whether to hide or show this Column
     *
     * @param boolean b - true to hide this Column, false to show
     */
    var isHidden: Boolean
        get() = myCol.isHidden
        set(b) {
            this.myCol.isHidden = b
        }

    init {
        wbh = mySheet.workBook
    }

    /**
     * resizes this column to fit the width of all displayed, non-wrapped text.
     * <br></br>NOTE: as the Excel autofit implementation is undocumented, this is an approximation
     */
    fun autoFit() {
        // KSC: make more betta :)
        var w = 0.0
        val cxt = this.cells
        for (t in cxt.indices) {
            val s = cxt[t].formattedStringVal    //StringVal();
            val fh = cxt[t].formatHandle
            val ef = fh!!.font
            var style = java.awt.Font.PLAIN
            if (ef!!.bold) style = style or java.awt.Font.BOLD
            if (ef.italic) style = style or java.awt.Font.ITALIC
            val h = ef.fontHeightInPoints.toInt()
            val f = java.awt.Font(ef.fontName, style, h)
            var newW = 0.0
            if (!cxt[t].formatHandle!!.wrapText)
            // normal case, no wrap
                newW = StringTool.getApproximateStringWidth(f, s)
            else
            // wrap - use current column width?????
                newW = (this.width / COL_UNITS_TO_PIXELS).toDouble()
            w = Math.max(w, newW)
            /*
	    int strlen = cstr.length();
	    int csz = strlen *= cxt[t].getFontSize();
	    int factor= 28;	// KSC: was 50 + added factor to guard below
	    if((csz*factor)>lastsz)
		this.setWidth(csz*factor);
		*/
        }
        if (w == 0.0)
            return     // keep original width ... that's what Excel does for blank columns ...
        // convert pixels to excel column units basically OpenXLS.COLUNITSTOPIXELS in double form
        this.width = Math.floor(w / DEFAULT_ZERO_CHAR_WIDTH * 256.0).toInt()
    }

    /**
     * sets the FormatHandle (a Format Object describing visual properties) for this Column
     * <br></br>NOTE: The Column format record describes the default formatting for each cell contained
     * within the column
     */
    private fun setFormatHandle() {
        if (formatter != null) return
        formatter = FormatHandle(wbh!!, this.formatId)
        formatter!!.setColHandle(this)
    }

    /**
     * determines if this Column passes through i.e. contains a
     * horizontal merge range
     *
     * @return true if this Column is part of any merge (horizontally merged cells)
     */
    fun containsMergeRange(): Boolean {
        val r = mySheet.rows
        for (i in r.indices) {
            val b: BiffRec?
            try {
                b = r[i].myRow.getCell(this.colFirst.toShort())
                if (b != null && b.mergeRange != null) return true
            } catch (e: CellNotFoundException) {
            }

        }
        return false
    }

    companion object {

        // TODO: read 1st font in file to set DEFAULT_ZERO_CHAR_WIDTH ... eventually ...
        val DEFAULT_ZERO_CHAR_WIDTH = 7.0 // width of '0' char in default font + conversion 1.3
        val COL_UNITS_TO_PIXELS = (256 / DEFAULT_ZERO_CHAR_WIDTH).toInt()  // = 36.57
        val DEFAULT_COLWIDTH = Colinfo.DEFAULT_COLWIDTH

        /**
         * static utility method to return the Column width of an existing column
         * in the units as follows:
         * <br></br>
         * default width of the columns in 1/256 of the width of the zero character,
         * using default font.
         * <br></br>For Arial 10 point, the default width of the zero character = 7
         * <br></br>The Default Excel Column, whose width in Characters or Excel Units is 8.43, has a width in these units of 2300.
         *
         * @param Boundsheet sheet - source Worksheet
         * @param int        col - 0-based Column number
         * @return int - Column width in internal units
         */
        fun getWidth(sheet: Boundsheet, col: Int): Int {
            var w = Colinfo.DEFAULT_COLWIDTH
            try {
                val c = sheet.getColInfo(col)
                if (c != null)
                    w = c.colWidth
            } catch (e: Exception) {    // exception if no col defined
            }

            return w
        }
    }
}