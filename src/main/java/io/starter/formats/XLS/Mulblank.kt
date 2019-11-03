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

import io.starter.OpenXLS.CellRange
import io.starter.OpenXLS.ExcelTools
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.ArrayList


/**
 * **Mulblank: Multiple Blank Cells (BEh)**<br></br>
 * This record stores up to 256 BLANK equivalents in
 * a space-saving format.
 *
 *
 * TODO: check compatibility with Excel2007 MAXCOLS
 *
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2       Row Number
 * 6       colFirst    2       Column Number of the first col of multiple Blank record
 * 8       rgixfe      var     Array of indexes to XF records
 * 10      colLast     2       Last Column containing Blank objects
</pre> *
 *
 * @see Blank
 */

class Mulblank : XLSCellRecord /*implements Mul*/() {

    internal var colFirst: Short = 0
    internal var colLast: Short = 0 // the colFirst/ColLast indexes determine
    internal var rgixfe: ByteArray? = null

    /**
     * since this is a "MUL" we override this method to
     * get a BiffRec Range, not a BiffRec Address.
     */
    override// KSC: if not referring to a single cell
    // referring to a single cell
    val cellAddress: String
        get() {
            var retval = "00"
            if (colNumber.toInt() == -1) {
                val rownum = rowNumber + 1
                retval = ExcelTools.getAlphaVal(colFirst.toInt()) + rownum
                retval += ":" + ExcelTools.getAlphaVal(colLast.toInt()) + rownum
            } else {
                val rownum = rowNumber + 1
                retval = ExcelTools.getAlphaVal(colNumber.toInt()) + rownum
            }
            return retval
        }

    /**
     * returns the cell address in int[] {row, col} format
     */
    override val intLocation: IntArray
        get() = if (colNumber.toInt() == -1) {
            intArrayOf(rw, colFirst.toInt(), rw, colLast.toInt())
        } else {
            intArrayOf(rw, colNumber.toInt())
        }


    /**
     * return entire range this Mulblank refers to
     *
     * @return
     */
    val mulblankRange: String
        get() {
            var retval = "00"
            val rownum = rowNumber + 1
            retval = ExcelTools.getAlphaVal(colFirst.toInt()) + rownum
            retval += ":" + ExcelTools.getAlphaVal(colLast.toInt()) + rownum
            return retval
        }

    /**
     * returnt the "current" column indicator, if set
     */
    override var colNumber: Short
        get() = if (colNumber.toInt() != -1) colNumber else colFirst


    /**
     * return a blank string val
     */
    override var stringVal: String
        get() = ""
        set

    /**
     * NOTE: Mublanks can have a portion of it's blank range which is merged: must determine if
     * the current cell is truly part of the merge range ...
     *
     * @return
     */
    override// this shouldn't happen ...
    // desired cell is NOT contained within master merge range
    var mergeRange: CellRange?
        get() {
            if (mergeRange == null) return null
            if (colNumber.toInt() == -1)
                return mergeRange
            return if (mergeRange!!.contains(intArrayOf(rowNumber, colNumber.toInt(), rowNumber, colNumber.toInt()))) mergeRange else null
        }
        set

    /**
     * get the ixfe for the desired referred-to blank
     */
    /**
     * sets the ixfe for the specific cell of the Mulblank
     * (each cell in a series of multiple blanks has their own ixfe)
     */
    override// set myxf to correct xf for cell in group of mulblanks
    var ixfe: Int
        get() {
            var idx = 0
            if (colNumber.toInt() != -1 && colNumber >= colFirst && colNumber <= colLast)
                idx = (colNumber - colFirst) * 2
            ixfe = ByteTools.readShort(rgixfe!![idx].toInt(), rgixfe!![idx + 1].toInt()).toInt()
            myxf = this.workBook!!.getXf(ixfe)
            return this.ixfe
        }
        set(i) {
            var idx = 0
            if (colNumber.toInt() != -1 && colNumber >= colFirst && colNumber <= colLast)
                idx = (colNumber - colFirst) * 2

            val b = ByteTools.shortToLEBytes(i.toShort())
            rgixfe[idx] = b[0]
            rgixfe[idx + 1] = b[1]
            updateRecord()
            ixfe = i
            myxf = this.workBook!!.getXf(ixfe)
        }

    /**
     * Get the referenced columns this mulblank has,
     */
    val colReferences: ArrayList<Int>
        get() {
            val colRefs = ArrayList<Int>()
            for (i in this.colFirst..this.colLast) {
                colRefs.add(i)
            }
            return colRefs
        }

    /**
     * returns the number of ixfe fields
     */
    internal val numFields: Int
        get() = colLast - colFirst + 1

    override fun toString(): String {
        return this.cellAddress
    }


    /**
     * set the Boundsheet for the Mulblank
     * this is needed because Blanks are BiffRec valrecs and
     * need to be assigned a BiffRec in the sheet...
     *
     *
     * the Mulblank itself does not get a cell.
     */
    override fun setSheet(bs: Sheet?) {
        this.worksheet = bs
    }

    /**
     * initialize this record
     */
    override fun init() {
        data = getData()
        super.init()
        if (this.length - 4 <= 0) {
            if (DEBUGLEVEL > -1)
                Logger.logInfo("no data in MULBLANK")
        } else {
            rw = ByteTools.readUnsignedShort(this.getByteAt(0), this.getByteAt(1))
            colFirst = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
            //col = colFirst;
            colNumber = -1    // flag that this rec hasn't been referred to one cell
            colLast = ByteTools.readShort(
                    this.getByteAt(this.reclen - 2).toInt(),
                    this.getByteAt(this.reclen - 1).toInt())
            //			Sometimes colFirst & colLast are reversed... WTFM$? -jm
            if (colLast < colFirst) {
                val csav = colLast
                colLast = colFirst
                colFirst = csav
                colLast--
            }
            if (DEBUGLEVEL > 5)
                Logger.logInfo(
                        "INFO: MULBLANK range: $colFirst:$colLast")
            val numblanks = colLast - colFirst + 1
            //			blanks = new ArrayList();
            if (numblanks < 1) {
                Logger.logWarn(
                        "WARNING: could not parse Mulblank record: numblanks reported  as:$numblanks")
                //Logger.logInfo((numblanks >> 12)*-1); ha!
                return
            }
            rgixfe = this.getBytesAt(4, numblanks * 2)
        }
        // KSC: to use as a blank:
        this.isValueForCell = true
        this.isBlank = true
    }

    /**
     * reset the "current" column use to reference a single blank of this Mulblank range of blank cells
     *
     * @return
     */
    fun resetCol() {
        colNumber = colFirst
    }

    /**
     * sets the first column of the range of blank cells referenced by this Mulblank
     *
     * @param c
     */
    fun setColFirst(c: Int) {
        colFirst = c.toShort()
    }

    /**
     * sets the last column of the range of blank cells referenced by this Mulblank
     *
     * @param c
     */
    fun setColLast(c: Int) {
        colLast = c.toShort()
    }

    /**
     * return sthe first column of the range of blank cells referenced by this Mulblank
     *
     * @return
     */
    override fun getColFirst(): Int {
        return colFirst.toInt()
    }

    /**
     * return sthe last column of the range of blank cells referenced by this Mulblank
     *
     * @return
     */
    override fun getColLast(): Int {
        return colLast.toInt()
    }

    /**
     * revise range of cells this Mulblank refers to; return true if no more blanks in range
     *
     * @param c col number to remove, 0-based
     */
    fun removeCell(c: Short): Boolean {
        if (c == colFirst) {
            colFirst++
            val tmp = ByteArray(rgixfe!!.size - 2)
            System.arraycopy(rgixfe!!, 2, tmp, 0, tmp.size)    // skip first
            rgixfe = tmp
        } else if (c == colLast) {
            colLast--
            val tmp = ByteArray(rgixfe!!.size - 2)
            System.arraycopy(rgixfe!!, 0, tmp, 0, tmp.size)    // skip last
            rgixfe = tmp
        }
        if (c > colFirst && c < colLast) {
            // must break apart Mulblank as now is non-contiguous ...
            // keep first colFirst->c as a MulBlank
            try {
                // create the blank records
                for (i in c + 1..colLast) {
                    val newblank = byteArrayOf(0, 0, 0, 0, 0, 0)
                    // set the row...
                    System.arraycopy(this.getBytesAt(0, 2)!!, 0, newblank, 0, 2)
                    // set the col...
                    System.arraycopy(ByteTools.shortToLEBytes(i.toShort()), 0, newblank, 2, 2)
                    // set the ixfe
                    System.arraycopy(rgixfe!!, (i - colFirst) * 2, newblank, 4, 2)
                    val b = Blank(newblank)
                    b.streamer = this.streamer
                    b.workBook = this.workBook
                    b.setSheet(this.sheet)
                    b.mergeRange = this.getMergeRange(i - colFirst)
                    this.row.removeCell(i.toShort())// remove this mulblank from the cells array
                    this.workBook!!.addRecord(b, true)    // and add a blank in it's place
                }
                // truncate the rgixfe:
                val tmp = ByteArray(2 * (c - colFirst + 1))
                System.arraycopy(rgixfe!!, 0, tmp, 0, tmp.size)    // skip last
                rgixfe = tmp
                // now truncate the Mulblank
                colLast = (c - 1).toShort()
            } catch (e: Exception) {
                Logger.logInfo("initializing Mulblank failed: $e")
            }

            colNumber = c    // the blank to remove
        }
        if (colFirst < 0 || colLast < 0) {    // can happen when removing multiple cells ..?
            return true
        }
        if (colFirst == colLast) {// covert to a single blank
            val newblank = byteArrayOf(0, 0, 0, 0, 0, 0)
            // set the row...
            System.arraycopy(this.getBytesAt(0, 2)!!, 0, newblank, 0, 2)
            // set the col...
            System.arraycopy(ByteTools.shortToLEBytes(colFirst), 0, newblank, 2, 2)
            // set the ixfe
            System.arraycopy(rgixfe!!, 0, newblank, 4, 2)
            val b = Blank(newblank)
            b.streamer = this.streamer
            b.workBook = this.workBook
            b.setSheet(this.sheet)
            b.mergeRange = this.getMergeRange(colFirst.toInt())
            colNumber = colFirst
            this.row.removeCell(this)// remove this mulblank from the cells array
            this.workBook!!.addRecord(b, true)
            colNumber = c    // still have to remove cell at col c
            return false    // no more mulblanks
        }
        updateRecord()
        // no more blanks in range ... can happen??
        return colFirst > colLast    // can delete it
        // don't delete this rec
    }

    /**
     * used to set the cell which this will be referred to, used when trying to access
     * ixfe
     *
     * @param c
     */
    fun setCurrentCell(c: Short) {
        colNumber = c
    }

    /**
     * retrieves the merged range for the desired cell in this group of blanks
     *
     * @param col
     * @return
     */
    private fun getMergeRange(col: Int): CellRange? {
        if (mergeRange == null) return null
        if (col == -1)
            return mergeRange    // this shouldn't happen ...
        return if (mergeRange!!.contains(intArrayOf(rowNumber, col, rowNumber, col))) mergeRange else null
// desired cell is NOT contained within master merge range
    }

    /**
     * set the column
     */
    override fun setCol(i: Short) {
        colNumber = i
    }

    /**
     * given new info (colFirst, colLast and rgixfe) update data record
     */
    private fun updateRecord() {
        val data = ByteArray(2 + 2 + 2 + rgixfe!!.size)
        data[0] = this.getData()!![0]        // row shouldn't have changed
        data[1] = this.getData()!![1]
        var b = ByteTools.shortToLEBytes(colFirst)
        data[2] = b[0]
        data[3] = b[1]
        // after colfirst= rgixfe
        System.arraycopy(rgixfe!!, 0, data, 4, rgixfe!!.size)
        b = ByteTools.shortToLEBytes(colLast)
        data[4 + rgixfe!!.size] = b[0]
        data[5 + rgixfe!!.size] = b[1]
        setData(data)
    }

    companion object {

        private val serialVersionUID = 2707362447402042745L

        val prototype: XLSRecord?
            get() {
                val mb = Mulblank()
                mb.opcode = XLSConstants.MULBLANK
                mb.setData(byteArrayOf(0, 0, 0, 0, 0, 0))
                mb.colNumber = -1
                return mb
            }
    }


}