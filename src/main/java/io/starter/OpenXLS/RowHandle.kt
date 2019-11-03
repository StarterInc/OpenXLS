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
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.json.JSONArray
import org.json.JSONException
import org.json.JSONObject

import java.awt.*
import java.util.ArrayList

import io.starter.OpenXLS.JSONConstants.*


/**
 * The RowHandle provides access to a Worksheet Row and its Cells.
 * <br></br>
 * Use the RowHandle to work with individual Rows in an XLS file.
 * <br></br>
 * With a RowHandle you can:
 * <br></br><blockquote>
 * get a handle to the Cells in a row<br></br>
 * set the default formatting for a Row<br></br>
 * <br></br></blockquote>
 *
 * @see WorkBookHandle
 *
 * @see WorkSheetHandle
 *
 * @see FormulaHandle
 */
class RowHandle(var myRow: Row, val workSheetHandle: WorkSheetHandle) {
    private var formatter: FormatHandle? = null
    val workBook: WorkBook?

    /**
     * Return the row height of an existing row.
     *
     *
     * These values are returned in twips, 1/20th of a character.
     *
     * @return int Height of Row in twips
     */
    /**
     * Sets the row height in twips (1/20th of a point)
     *
     * @param newHeight
     */
    /* 20080604 KSC: if an image falls upon this column,
         * adjust image width so that it does not change
         */// for each image that falls over this row, trap index + original width -- to be reset after setting row height
    var height: Int
        get() = myRow.rowHeight
        set(newHeight) {
            val iAdjust = ArrayList<IntArray>()
            val images = myRow.sheet!!.images
            if (images != null) {
                for (z in images.indices) {
                    val ih = images[z]
                    val r0 = ih.row
                    val r1 = ih.row1
                    val row = myRow.rowNumber
                    if (row >= r0 && row <= r1) {
                        val h = ih.height.toInt()
                        iAdjust.add(intArrayOf(z, h))
                    }
                }
            }
            myRow.rowHeight = newHeight
            for (z in iAdjust.indices) {
                val ih = images!![iAdjust[z][0]]
                ih.setHeight(iAdjust[z][1])
            }

        }

    /**
     * returns the row height in Excel units, which depends upon the default font
     * <br></br>in Arial 10 pt, standard row height is 12.75 points
     *
     * @return int row height in Excel units
     */
    /**
     * sets the row height in Excel units.
     *
     * @param double i - row height value in Excel units
     */
    // 20090506 KSC: apparently it's in twips ?? 1/20 of a point
    var heightInChars: Int
        get() = myRow.rowHeight / 20
        set(newHeight) {
            this.height = newHeight * 20
        }

    /**
     * Gets the FormatHandle for this Row.
     *
     * @return FormatHandle - a Format object to apply to this Row
     */
    val formatHandle: FormatHandle?
        get() {
            if (this.formatter == null) this.setFormatHandle()
            return this.formatter
        }

    /**
     * gets the current default row format id.  May be overwritten by contained cells
     *
     * @return format id of row
     */
    /**
     * sets the default format id for the Row's Cells
     *
     * @param int Format Id for all Cells in Row
     */
    var formatId: Int
        get() = if (myRow.explicitFormatSet) myRow.ixfe else this.workBook!!.workBook.defaultIxfe
        set(i) {
            myRow.ixfe = i
        }

    /**
     * Returns the array of Cells in this Row
     *
     * @return Cell[] all Cells in this Row
     */
    // don't use cache
    val cells: Array<CellHandle>
        get() = getCells(false)


    /**
     * Get the JSON object for this row.
     *
     * @return
     */
    val json: String
        get() = getJSON(255).toString()

    /**
     * Returns the row number of this RowHandle
     */
    val rowNumber: Int
        get() = myRow.rowNumber

    /**
     * Returns the Outline level (depth) of the row
     *
     * @return
     */
    /**
     * Set the Outline level (depth) of the row
     *
     * @param x
     */
    var outlineLevel: Int
        get() = myRow.outlineLevel
        set(x) {
            myRow.outlineLevel = x
        }

    /**
     * Returns whether the row is collapsed
     *
     * @return
     */
    /**
     * Set whether the row is collapsed.
     * Will hide the current row, and all contiguous rows
     * with the same outline level.
     *
     * @param b
     */
    var isCollapsed: Boolean
        get() = myRow.isCollapsed
        set(b) {
            myRow.isCollapsed = b
        }

    /**
     * Returns whether the row is hidden
     *
     * @return
     */
    /**
     * Set whether the row is hidden
     *
     * @param b
     */
    var isHidden: Boolean
        get() = myRow.isHidden
        set(b) {
            myRow.isHidden = b
        }


    /**
     * true if row height has been altered from default
     * i.e. set manually
     *
     * @return
     */
    val isAlteredHeight: Boolean
        get() = myRow.isAlteredHeight

    /**
     * returns true if there is a Thick Top border set on the row
     */
    /**
     * sets this row to have a thick top border
     */
    var hasThickTopBorder: Boolean
        get() = myRow.hasThickTopBorder
        set(hasBorder) {
            myRow.hasThickTopBorder = hasBorder
        }

    /**
     * returns true if there is a Thick Bottom border set on the row
     */
    /**
     * sets this row to have a thick bottom border
     */
    var hasThickBottomBorder: Boolean
        get() = myRow.hasThickBottomBorder
        set(hasBorder) {
            myRow.hasThickBottomBorder = hasBorder
        }

    /**
     * returns true if there is a thick top or thick or medium bottom border on previoous row
     *
     *
     * Not useful for public API
     */
    val hasAnyThickTopBorder: Boolean
        get() = myRow.hasAnyThickTopBorder

    /**
     * Additional space below the row. This flag is set, if the
     * lower border of at least one cell in this row or if the upper
     * border of at least one cell in the row below is formatted with
     * a medium or thick line style. Thin line styles are not taken
     * into account.
     *
     *
     * Usage of this method is primarily for UI applications, and is not
     * needed for standard OpenXLS functionality
     */
    val hasAnyBottomBorder: Boolean
        get() = myRow.hasAnyBottomBorder

    /**
     * return the min/max columns defined for this row
     *
     * @return
     */
    val colDimensions: IntArray
        get() = myRow.colDimensions

    init {
        workBook = workSheetHandle.workBook
    }

    /**
     * sets the row height to auto fit
     * <br></br>When the row height is set manually, autofit is automatically turned off
     */
    fun setRowHeightAutoFit() {
        // this.myRow.setUnsynched(false);	// firstly, set so excel
        val ct = myRow.cells
        val it = ct.iterator()
        var h = 0.0
        val dpi = Toolkit.getDefaultToolkit().screenResolution        // 96 is "small", 120 dpi is "lg"
        // 1 point= 1/72 of an inch
        // 1 twip=  1 twip= 1/20 of a point
        // this should be a pretty good pixels/twips conversion factor.
        val factorTwip = dpi.toDouble() / 72.0 / 20.0        // .06 is "normal"
        // factorZero is width of 0 char in default font.  If assume Arial 10 pt, it is 6 + 1= 7
        val factorZero = 7.0    //java.awt.Toolkit.getDefaultToolkit().getFontMetrics(f).charWidth('0') + 1;

        while (it.hasNext()) {
            val cellrec = it.next() as XLSRecord
            try {
                var newH = 255.0    // default row height
                try {
                    val ef = cellrec.xfRec!!.font
                    var style = java.awt.Font.PLAIN
                    if (ef!!.bold)
                        style = style or java.awt.Font.BOLD
                    if (ef.italic)
                        style = style or java.awt.Font.ITALIC
                    val f = java.awt.Font(ef.fontName, style, ef.fontHeightInPoints.toInt())
                    val s = cellrec.stringVal
                    if (!cellrec.xfRec!!.wrapText)
                    // normal case, no wrap
                        newH = StringTool.getApproximateHeight(f, s, java.lang.Double.MAX_VALUE)
                    else {                // wrap to column width
                        // convert column width to pixels
                        // factorZero is usually 7		// double factorZero= java.awt.Toolkit.getDefaultToolkit().getFontMetrics(f).charWidth('0') + 1;
                        val cw = ColHandle.getWidth(this.workSheetHandle.boundsheet, cellrec.colNumber.toInt()) / 256.0
                        newH = StringTool.getApproximateHeight(f, s, cw * factorZero)
                    }
                    // this doesn't work correctly		    newH/=factorTwip;	// pixels * twips/pixels == twips)
                    newH *= 20.0    // this is better ...
                } catch (e: Exception) {
                    Logger.logErr("RowHandle.setRowHeightAutoFit: $e")
                }

                h = Math.max(h, newH)
            } catch (e: Exception) {
            }

        }
        if (h > 0)
            this.myRow.rowHeight = Math.ceil(h).toInt()
    }

    /**
     * Determines if the row passes through
     * a vertical merge range
     *
     * @return
     */
    fun containsVerticalMergeRange(): Boolean {
        val c = this.cells
        for (i in c.indices) {
            if (c[i].mergedCellRange != null) {
                val cr = c[i].mergedCellRange
                try {
                    if (cr!!.rows.size > 1) return true
                } catch (e: Exception) {
                }

            }
        }
        return false
    }


    /**
     * Set up the format handle for this row
     */
    private fun setFormatHandle() {
        if (formatter != null) return
        formatter = FormatHandle(workBook!!, this.formatId)
        formatter!!.setRowHandle(this)
    }

    /**
     * Returns the array of Cells in this Row
     *
     * @param cache cellhandles flag
     * @return Cell[] all Cells in this Row
     */
    fun getCells(cached: Boolean): Array<CellHandle> {
        val ct = myRow.cells
        val it = ct.iterator()
        val ch = arrayOfNulls<CellHandle>(ct.size)
        var t = 0
        var aMul: Mulblank? = null
        var c: Short = -1
        while (it.hasNext()) {
            val rc = it.next() as BiffRec
            try {  // use cache of Cellhandles!
                if (rc.opcode != XLSConstants.MULBLANK) {
                    ch[t] = this.workSheetHandle.getCell(rc.rowNumber, rc.colNumber.toInt(), cached)
                } else {
                    // handle Mulblanks: ref a range of cells; to get correct cell address,
                    // traverse thru range and set cellhandle ref to correct column
                    if (rc === aMul) {
                        c++
                    } else {
                        aMul = rc as Mulblank
                        c = aMul.colFirst.toShort()
                    }
                    ch[t] = this.workSheetHandle.getCell(rc.rowNumber, c.toInt(), cached)
                }
            } catch (cnfe: CellNotFoundException) {
                rc.setXFRecord()
                ch[t] = CellHandle(rc, null)
                ch[t].workSheetHandle = null //TODO: implement if causing grief -jm
                if (rc.opcode == XLSConstants.MULBLANK) {
                    // handle Mulblanks: ref a range of cells; to get correct cell address,
                    // traverse thru range and set cellhandle ref to correct column
                    if (rc === aMul) {
                        c++
                    } else {
                        aMul = rc as Mulblank
                        c = aMul.colFirst.toShort()
                    }
                    ch[t].setBlankRef(c.toInt())    // for Mulblank use only -sets correct column reference for multiple blank cells ...
                }

            }

            t++
        }
        return ch
    }

    fun getJSON(maxcols: Int): JSONObject {
        val theRange = JSONObject()
        val cells = JSONArray()
        try {
            theRange.put(JSON_ROW, rowNumber)

            theRange.put(JSON_ROW_BORDER_TOP, hasAnyThickTopBorder)
            theRange.put(JSON_ROW_BORDER_BOTTOM, hasAnyBottomBorder)
            if (formatId != workBook!!.workBook.defaultIxfe) theRange.put("xf", formatId)
            theRange.put(JSON_HEIGHT, height / ROW_HEIGHT_DIVISOR + 5) // the default is TOO SMALL!
            val chandles = getCells(false)
            var i = 0
            while (i < chandles.size) {
                val thisCell = chandles[i]
                if (!thisCell.isDefaultCell) {
                    // do NOT use cached formula vals
                    if (thisCell.cell!!.opcode == XLSRecord.FORMULA) {
                        try {
                            val fh = thisCell.formulaHandle
                            fh.formulaRec!!.setCachedValue(null)
                        } catch (ex: Exception) {
                        }

                    }

                    if (thisCell.colNum >= maxcols) {
                        i = chandles.size
                    } else if (thisCell.cell!!.opcode == XLSRecord.MULBLANK) {
                        val mb = thisCell.cell as Mulblank
                        val columns = mb.colReferences
                        for (x in columns.indices) {
                            thisCell.setBlankRef(columns[x])
                            thisCell.cell!!.setCol(columns[x].toShort())
                            val result = JSONObject()
                            thisCell.cellAddress
                            var v: Any? = ""
                            try {
                                v = thisCell.jsonObject
                            } catch (exz: Exception) {
                                Logger.logErr("Error getting Row cell value " + thisCell.cellAddress + " JSON: " + exz)
                                v = "ERROR FETCHING VALUE for:" + thisCell.cellAddress
                            }

                            if (v != null) {
                                result.put(JSON_CELL, v)
                                cells.put(result)
                            }
                        }
                    } else {
                        val result = JSONObject()
                        var v: Any? = "ERROR FETCHING VALUE for:" + thisCell.cellAddress
                        try {
                            v = thisCell.jsonObject
                        } catch (exz: Exception) {
                            Logger.logErr("Error getting Row cell value " + thisCell.cellAddress + " JSON: " + exz)
                        }

                        if (v != null) {
                            result.put(JSON_CELL, v)
                            cells.put(result)
                        }
                    }
                }
                i++
            }
            theRange.put(JSON_CELLS, cells)
        } catch (e: JSONException) {
            Logger.logErr("Error getting Row JSON: $e")
        }

        return theRange
    }

    /**
     * Returns the String representation of this Row
     */
    override fun toString(): String {
        return myRow.toString()
    }

    fun setBackgroundColor(colr: java.awt.Color) {
        setFormatHandle()
        formatter!!.setCellBackgroundColor(colr)
    }

    companion object {
        // FYI: do not change lightly -- these match Excel 2007 almost exactly
        var ROW_HEIGHT_DIVISOR = 17


        /**
         * Return the row height of an existing row.
         *
         *
         * These values are returned in twips, 1/20th of a character.
         *
         * @param sheet
         * @param row
         * @return
         */
        fun getHeight(sheet: Boundsheet, row: Int): Int {
            var h = 255
            try {
                val r = sheet.getRowByNumber(row)
                if (r != null)
                    h = r.rowHeight
            } catch (e: Exception) {    // exception if no row defined
                h = 255 // default
            }

            return h
        }
    }
}
