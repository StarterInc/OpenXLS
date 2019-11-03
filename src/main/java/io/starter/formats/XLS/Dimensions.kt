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
package io.starter.formats.XLS

import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger


/**
 * **Dimensions 0x200: Describes the max BiffRec dimensions of a Sheet.**<br></br>
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       rowMic      4       First defined Row
 * 8       rowMac      4       Last defined Row plus 1
 * 12      colFirst    2       First defined column
 * 14      colLast     2       Last defined column plus 1
 * 16      reserved    2       must be 0
 *
 * When a record is added to a Boundsheet, we need to check if this
 * has changed the row/col dimensions of the sheet and update this
 * record accordingly.
 *
</pre> *
 *
 * @see WorkBook
 *
 * @see Boundsheet
 *
 * @see Index
 *
 * @see Dbcell
 *
 * @see Row
 *
 * @see Cell
 *
 * @see XLSRecord
 */


class Dimensions : io.starter.formats.XLS.XLSRecord() {
    internal var rowFirst = 0
    internal var rowLast = 0
    internal var colFirst: Short = 0
    internal var colLast: Short = 0


    /**
     * set last/first cols/rows
     */
    fun setRowFirst(c: Int) {
        var c = c
        c++ // inc here instead of updateRowDimensions
        val b = ByteTools.cLongToLEBytes(c)
        val dt = this.getData()
        if (dt!!.size > 4) System.arraycopy(b, 0, dt, 0, 4)
        this.rowFirst = c
    }

    fun getRowFirst(): Int {
        return rowFirst
    }

    /**
     * set last/first cols/rows
     */
    fun setRowLast(c: Int) {
        var c = c
        c++ // inc here instead of updateRowDimensions
        val b = ByteTools.cLongToLEBytes(c)
        val dt = this.getData()
        if (dt!!.size > 4) System.arraycopy(b, 0, dt, 4, 4)
        this.rowLast = c
    }

    fun getRowLast(): Int {
        return rowLast
    }

    /**
     * set last/first cols/rows
     */
    fun setColFirst(c: Int) {
        val b = ByteTools.shortToLEBytes(c.toShort())
        val dt = this.getData()
        if (dt!!.size > 4) System.arraycopy(b, 0, dt, 8, 2)
        this.colFirst = c.toShort()
    }

    fun getColFirst(): Int {
        return colFirst.toInt()
    }

    /**
     * set last/first cols/rows
     */
    fun setColLast(c: Int) {
        var c = c
        c++
        if (c >= XLSConstants.MAXCOLS_BIFF8 && // warn about maxcols
                !this.wkbook!!.isExcel2007)
            Logger.logWarn("Dimensions.setColLast column: $c is incompatible with pre Excel2007 versions.")

        if (c >= XLSConstants.MAXCOLS)
            c = XLSConstants.MAXCOLS// odd case, its supposed to be last defined col +1, but this breaks last col
        val b = ByteTools.shortToLEBytes(c.toShort())
        val dt = this.getData()
        if (dt!!.size > 4) System.arraycopy(b, 0, dt, 10, 2)
        this.colLast = c.toShort()
    }

    fun getColLast(): Int {
        return colLast.toInt()
    }

    override fun setSheet(bs: Sheet?) {
        super.setSheet(bs)
        bs!!.dimensions = this
    }

    override fun init() {
        super.init()
        rowFirst = ByteTools.readInt(this.getByteAt(0), this.getByteAt(1), this.getByteAt(2), this.getByteAt(3))
        rowLast = ByteTools.readInt(this.getByteAt(4), this.getByteAt(5), this.getByteAt(6), this.getByteAt(7))

        colFirst = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())
        colLast = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt())
        this.getData()
    }

    /**
     * update the min/max cols and rows
     */
    fun updateDimensions(row: Int, col: Short) {
        this.updateRowDimensions(row)
        this.updateColDimension(col)
    }

    /**
     * update the min/max cols and rows
     */
    fun updateRowDimensions(row: Int) {
        // check row dimension
        //row++; // TODO: check why we are incrementing here... nowincremented in setRowLast
        if (row >= rowLast) {
            this.setRowLast(row) // now incremented only in setRowXX
        }
        if (row < rowFirst) {
            this.setRowFirst(row) // now incremented only in setRowXX
        }

        if (DEBUGLEVEL > 10) {
            val shtnm = this.sheet!!.sheetName
            Logger.logInfo("$shtnm dimensions: $rowFirst:$colFirst-$rowLast:$colLast")
        }
    }

    /**
     * update the min/max cols and rows
     */
    fun updateColDimension(col: Short) {
        // check cell dimension
        if (col > colLast) {
            this.setColLast(col.toInt())
        } else if (col - 1 < colFirst) {
            this.setColFirst(col.toInt())
        }
        if (DEBUGLEVEL > 10) {
            val shtnm = this.sheet!!.sheetName
            Logger.logInfo("$shtnm dimensions: $rowFirst:$colFirst-$rowLast:$colLast")
        }
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 7156425132146869228L
    }

}