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
package io.starter.formats.XLS.formulas

import io.starter.formats.XLS.Shrfmla
import io.starter.formats.XLS.XLSConstants
import io.starter.formats.XLS.XLSRecord


/**
 * ptgArea is a reference to an area (rectangle) of cells.
 * Essentially it is a collection of two ptgRef's, so it will be
 * treated that way in the code...
 *
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           rwFirst     	2       The First row of the reference
 * 2           rwLast     		2       The Last row of the reference
 * 4           grbitColFirst   2       (see following table)
 * 6           grbitColLast    2       (see following table)
 *
 * Only the low-order 14 bits specify the Col, the other bits specify
 * relative vs absolute for both the col or the row.
 *
 * Bits        Mask        Name    Contents
 * -----------------------------------------------------
 * 15          8000h       fRwRel  =1 if row offset relative,
 * =0 if otherwise
 * 14          4000h       fColRel =1 if row offset relative,
 * =0 if otherwise
 * 13-0        3FFFh       col     Ordinal column offset or number
</pre> *
 *
 * @see Ptg
 *
 * @see Formula
 */
class PtgAreaN : PtgArea() {

    override val isOperand: Boolean
        get() = true

    override val isReference: Boolean
        get() = true

    /**
     * @return firstPtgN
     */
    var firstPtgN: PtgRefN? = null
        internal set
    /**
     * @return lastPtgN
     */
    var lastPtgN: PtgRefN? = null
        internal set
    private var parea: PtgArea? = null

    /**
     * returns the row/col ints for the ref
     *
     * @return
     */
    override val rowCol: IntArray?
        get() {
            if (firstPtgN == null) {
                val rc1 = firstPtgN!!.rowCol
                return intArrayOf(rc1[0], rc1[1], rc1[0], rc1[1])
            }
            val rc1 = firstPtgN!!.rowCol
            val rc2 = lastPtgN!!.rowCol
            return intArrayOf(rc1[0], rc1[1], rc2[0], rc2[1])
        }

    /**
     * returns the uncoverted, actual row col
     *
     * @return
     */
    val realRowCol: IntArray
        get() = intArrayOf(firstPtgN!!.rw, firstPtgN!!.col, lastPtgN!!.rw, lastPtgN!!.col)

    /*
    Returns the location of the Ptg as a string
*/
    override// we tried
    var location: String?
        get() {
            if (firstPtgN == null || lastPtgN == null) {
                this.populateVals()
                if (firstPtgN == null || lastPtgN == null)
                    throw AssertionError("PtgAreaN.getLocation null ptgs")
            }
            val s = firstPtgN!!.location
            val y = lastPtgN!!.location

            return "$s:$y"
        }
        set

    /**
     * returns an array of the first and last addresses in the PtgAreaN
     */
    // 20060223: KSC: customize from ptgArea
    override val intLocation: IntArray?
        get() {
            val returning = IntArray(4)
            try {
                val first = firstPtgN!!.intLocation
                val last = lastPtgN!!.intLocation
                System.arraycopy(first!!, 0, returning, 0, 2)
                System.arraycopy(last!!, 0, returning, 2, 2)
            } catch (e: Exception) {
            }

            return returning
        }

    /**
     * custom RefTracker usage:  uses entire range covered by all shared formulas
     */
    // TODO: determine if this is an OK maxcol (Excel 2007)
    // TODO: determine if this is an OK maxcol (Excel 2007)
    val area: PtgArea
        get() {
            val sh = this.parentRec as Shrfmla
            val i = IntArray(4)
            if (fRwRel) {
                i[0] = sh.firstRow + firstPtgN!!.rw
            } else {
                i[0] = firstPtgN!!.rw
            }
            if (fColRel) {
                i[1] = sh.firstCol + firstPtgN!!.col
            } else {
                i[1] = firstPtgN!!.col
            }
            if (fRwRel) {
                i[2] = sh.lastRow + lastPtgN!!.rw
            } else {
                i[2] = lastPtgN!!.rw
            }
            if (fColRel) {
                i[3] = sh.lastCol + lastPtgN!!.col
            } else {
                i[3] = lastPtgN!!.col
            }

            if (i[1] >= XLSConstants.MAXCOLS_BIFF8 && !this.parentRec!!.workBook!!.isExcel2007)
                i[1] -= XLSConstants.MAXCOLS_BIFF8
            if (i[3] >= XLSConstants.MAXCOLS_BIFF8 && !this.parentRec!!.workBook!!.isExcel2007)
                i[3] -= XLSConstants.MAXCOLS_BIFF8

            return PtgArea(i, sh, true)
        }

    /*
     Throw this data into two ptgref's
    */
    override fun populateVals() {
        val temp1 = ByteArray(5)
        val temp2 = ByteArray(5)
        temp1[0] = 0x24
        temp2[0] = 0x24
        System.arraycopy(record, 1, temp1, 1, 2)
        System.arraycopy(record, 5, temp1, 3, 2)
        System.arraycopy(record, 3, temp2, 1, 2)
        System.arraycopy(record, 7, temp2, 3, 2)

        /*
 *  0           rwFirst     2       The First row of the reference
    2           rwLast     2       The Last row of the reference
    4           grbitColFirst    2       (see following table)
    6           grbitColLast    2       (see following table)

    Only the low-order 14 bits specify the Col, the other bits specify
    relative vs absolute for both the col or the row.

    Bits        Mask        Name    Contents
    -----------------------------------------------------
    15          8000h       fRwRel  =1 if row offset relative,
                                    =0 if otherwise
    14          4000h       fColRel =1 if row offset relative,
                                    =0 if otherwise
    13-0        3FFFh       col     Ordinal column offset or number

 */
        firstPtgN = PtgRefN(false)
        firstPtgN!!.parentRec = parentRec
        firstPtgN!!.init(temp1)
        lastPtgN = PtgRefN(false)
        lastPtgN!!.parentRec = parentRec
        lastPtgN!!.init(temp2)
        if (parentRec != null && parentRec is Shrfmla) {
            // 20060301 KSC: init sets formula row to 1st row for a shared formula; adjust here
            lastPtgN!!.setFormulaRow((parentRec as Shrfmla).lastRow)
            lastPtgN!!.setFormulaCol((parentRec as Shrfmla).lastCol)
        }
        //        if (this.useReferenceTracker)
        //        	this.addToRefTracker();
    }

    fun convertToPtgArea(r: io.starter.formats.XLS.XLSRecord): PtgArea {
        val p1 = firstPtgN!!.convertToPtgRef(r)
        val p2 = lastPtgN!!.convertToPtgRef(r)
        return PtgArea(p1, p2, r)
    }


    /**
     * update record bytes
     */
    // 20060223 KSC
    override fun updateRecord() {
        val first = firstPtgN!!.record
        val last = lastPtgN!!.record
        // the last record has an extra identifier on it.
        val newrecord = ByteArray(9)
        newrecord[0] = record[0]
        System.arraycopy(first, 1, newrecord, 1, 2)
        System.arraycopy(last, 1, newrecord, 3, 2)
        System.arraycopy(first, 3, newrecord, 5, 2)
        System.arraycopy(last, 3, newrecord, 7, 2)
        record = newrecord
    }

    /**
     * add "true" area to reference tracker i.e. entire range referenced by all shared formula members
     */
    override fun addToRefTracker() {
        val iParent = this.parentRec!!.opcode.toInt()
        if (iParent == XLSConstants.SHRFMLA.toInt()) {
            // KSC: TESTING - local ptgarea gets finalized and messes up ref. tracker on multiple usages without close
            //getArea();
            //parea.addToRefTracker();
            val parea = this.area    // is finalized if local var --- but take out ptgarea finalize for now
            parea.addToRefTracker()
        }
    }

    /**
     * remove "true" area from reference tracker i.e. entire range referenced by all shared formula members
     */
    override fun removeFromRefTracker() {
        val iParent = this.parentRec!!.opcode.toInt()
        if (iParent == XLSConstants.SHRFMLA.toInt()) {
            val parea = this.area
            parea.removeFromRefTracker()
        }
        //if (parea!=null) {
        //	parea.removeFromRefTracker();
        //parea.close();
        //}
        //parea= null;
    }

    override fun close() {
        removeFromRefTracker()
        if (parea != null)
            parea!!.close()
        parea = null
        if (firstPtgN != null)
            firstPtgN!!.close()
        firstPtgN = null
        if (lastPtgN != null)
            lastPtgN!!.close()
        lastPtgN = null
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -8433468704529379504L
    }
}	
