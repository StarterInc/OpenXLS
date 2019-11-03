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

import io.starter.OpenXLS.ExcelTools
import io.starter.formats.XLS.formulas.*
import io.starter.toolkit.ByteTools

import java.util.*

/**
 * SHRFMLA is a file optimization that stores many identical formulas once.
 * The SHRFMLA record appears immediately after the first FORMULA record in the
 * group. Each FORMULA record in the group will have fShrFmla set and will
 * contain only a PtgExp pointing to the cell containing the SHRFMLA record.
 *
 *
 * Occasionally Excel writes a FORMULA record with fShrFmla set whose
 * expression is an instantiation of the relevant shared formula instead of the
 * usual PtgExp. We currently handle these by clearing fShrFmla. At some point
 * in the future we will attempt to detect which shared formula is applied and
 * make the cell a member of its group.
 *
 *
 * OpenXLS does not currently create shared formulas or add cells to existing
 * ones. Existing shared formulas will be preserved in output. Removal of
 * formulas from the group is supported, including the cell currently hosting
 * the SHRFMLA record. Shared formula member cells will be reference-tracked
 * and recalculated properly.
 * <pre>
 * OFFSET NAME        SIZE CONTENTS
 * 0      rwFirst     2    First Row
 * 2      rwLast      2    Last Row
 * 4      colFirst    1    First Column
 * 5      colLast     1    Last Column
 * 6      (reserved)  2    pass through, zero for new
 * 8      cce         2    Length of the parsed expression
 * 10     rgce        cce  Parsed Expression
</pre> *
 */
class Shrfmla : XLSRecord() {

    private var rwFirst: Int = 0
    var lastRow: Int = 0
    var firstCol: Int = 0
        private set
    var lastCol: Int = 0
        private set
    var stack: Stack<*>? = null
        private set
    private var ptgcache: Array<Ptg>? = null
    private var host: Formula? = null

    /**
     * The set of Formula records referring to this shared formula.
     */
    /**
     * return all the formulas that use this Shrfmla
     *
     * @return
     */
    var members: SortedSet<*>? = TreeSet(CellAddressComparator())
        private set

    /**
     * Whether this formula contains an indirect reference.
     */
    internal var containsIndirectFunction = false

    var firstRow: Int
        get() = rwFirst
        set(row) {
            rwFirst = row
            rw = rwFirst
        }

    /**
     * Gets the area where references to this shared formula may occur.
     *
     * @return an A1-style range string
     */
    val cellRange: String
        get() = ExcelTools.formatRange(
                intArrayOf(firstCol, rwFirst, lastCol, lastRow))

    /**
     * Gets the formula record in which this record resides.
     */
    /**
     * Sets the formula record in which this record resides.
     */
    var hostCell: Formula?
        get() = host
        set(newHost) {
            if (host != null) host!!.removeInternalRecord(this)

            host = newHost
            host!!.addInternalRecord(this)
            rw = host!!.rowNumber
            colNumber = host!!.colNumber
        }

    /**
     * Gets a [PtgExp] pointer to this `ShrFmla`.
     * The returned `PtgExp` points to this record at its current
     * location. If the host cell or its address changes the pointer will
     * become invalid.
     */
    val pointer: PtgExp
        get() {
            val pointer = PtgExp()
            pointer.parentRec = host
            pointer.init(host!!.rowNumber, host!!.colNumber.toInt())
            return pointer
        }

    /**
     * update location upon a shift (row insertion or deletion)-- ensure member formulas cache are cleared
     *
     * @param shiftamount
     */
    fun updateLocation(shiftamount: Int, pr: PtgRef) {
        // remove original reference
        if (ptgcache!!.size > 1) {
            // for shrfmlas which contain multiple ptgs, ensure formulas get shifted only 1x!
            if (ptgcache!![0] is PtgRefN) {
                if (pr.hashcode != (ptgcache!![0] as PtgRefN).area.hashcode)
                    return     // it's already been shifted
            } else {
                if (pr.hashcode != (ptgcache!![0] as PtgAreaN).area.hashcode)
                    return     // it's already been shifted
            }
        }
        for (i in ptgcache!!.indices) {
            if (ptgcache!![i] is PtgRefN)
                (ptgcache!![i] as PtgRefN).removeFromRefTracker()
            else
                (ptgcache!![i] as PtgAreaN).removeFromRefTracker()
        }
        val ii = members!!.iterator()
        while (ii.hasNext()) {
            val f = ii.next()
            f.clearCachedValue()
            // also update PtgExp
            val pointer = f.getExpression()!!.get(0) as PtgExp
            pointer.setRowFirst(pointer.rwFirst + shiftamount)
        }
        firstRow = rwFirst + shiftamount
        lastRow = lastRow + shiftamount
        for (i in ptgcache!!.indices) {
            if (ptgcache!![i] is PtgRefN)
                (ptgcache!![i] as PtgRefN).addToRefTracker()
            else
                (ptgcache!![i] as PtgAreaN).addToRefTracker()
        }
    }

    override fun init() {
        super.init()
        rwFirst = ByteTools.readUnsignedShort(this.getByteAt(0), this.getByteAt(1))
        lastRow = ByteTools.readUnsignedShort(this.getByteAt(2), this.getByteAt(3))
        firstCol = this.getByteAt(4).toInt()
        lastCol = this.getByteAt(5).toInt()
        val cce = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())
        val rgce = this.getBytesAt(10, cce.toInt())
        if (this.sheet == null) this.setSheet(this.wkbook!!.lastbound)
        rw = rwFirst

        try {
            wkbook!!.lastFormula!!.initSharedFormula(this)
            this.hostCell = wkbook!!.lastFormula
        } catch (e: Exception) {
        }

        stack = ExpressionParser.parseExpression(rgce, this)
        if (containsIndirectFunction) {
            if (host != null) host!!.registerIndirectFunction()
        }
        // Cache Relative Ptgs for quickness of access
        val ptgs = ArrayList()
        for (idx in stack!!.indices) {
            val ptg = stack!![idx] as Ptg
            if (ptg is PtgRefN) {
                ptgs.add(ptg)
            } else if (ptg is PtgAreaN) {
                ptgs.add(ptg)
            }
        }
        ptgcache = arrayOfNulls<Ptg>(ptgs.size)
        ptgs.toTypedArray()
    }

    override fun preStream() {
        super.preStream()

        val data = getData()

        System.arraycopy(
                ByteTools.shortToLEBytes(rwFirst.toShort()), 0,
                data!!, 0, 2)

        System.arraycopy(
                ByteTools.shortToLEBytes(lastRow.toShort()), 0,
                data, 2, 2)

        data[4] = firstCol.toByte()
        data[5] = lastCol.toByte()

        data[7] = members!!.size.toByte()

        setData(data)
    }

    internal fun isInRange(s: String): Boolean {
        return ExcelTools.isInRange(s, rwFirst, lastRow, firstCol, lastCol)
    }

    override fun toString(): String {
        return ("Shared Formula [" + cellRange + "] "
                + FormulaParser.getExpressionString(stack!!))
    }

    /**
     * Adds a member formula to this shared formula.
     *
     * @throws IndexOutOfBoundsException if there are already 255 members
     */
    fun addMember(member: Formula) {
        if (members!!.size >= 255)
            throw IndexOutOfBoundsException(
                    "shared formula already has 255 members")

        members!!.add(member)
        if (members!!.size == 1)
        // KSC: ADDED
            hostCell = member

        // Only do range/host manipulation if we're not parsing
        if (!workBook!!.factory!!.iscompleted()) return

        // If the newly added member is the first one it must become the host
        //    	if (members.first() == member) setHostCell( member );  KSC: replaced with above

        val row = member.rowNumber
        val col = member.colNumber.toInt()

        if (row < rwFirst) rwFirst = row
        if (row > lastRow) lastRow = row
        if (col < firstCol) firstCol = col
        if (col > lastCol) lastCol = col
    }

    fun removeMember(member: Formula) {
        members!!.remove(member)

        // If we've just removed the last member, don't bother adjusting
        // because we're about to be deleted.
        if (members!!.size == 0) return

        /* Ideally we would shrink the range to the smallest possible value,
         * but it's not actually required. Finding the column bounds is somewhat
         * expensive as it requires iterating the member list. Therefore we
         * only update the row components of the range.
         */

        if (member.rowNumber == lastRow)
            lastRow = (members!!.last() as Formula).rowNumber.toShort().toInt()

        // If we're removing the host cell, choose another one
        if (member == host) {
            hostCell = members!!.first() as Formula
            rwFirst = host!!.rowNumber
        }
    }

    /**
     * convert expression using dimensions of specific member formula
     *
     * @param parent
     * @return
     */
    fun instantiate(parent: Formula): Stack<*> {
        return convertStack(stack!!, parent)
    }

    /**
     * Set if the formula contains Indirect()
     *
     * @param containsIndirectFunction The containsIndirectFunction to set.
     */
    fun setContainsIndirectFunction(containsIndirectFunction: Boolean) {
        this.containsIndirectFunction = containsIndirectFunction
    }

    /**
     * Adds an indirect function to the list of functions to be evaluated post load
     *
     * *
     * NOT IMPLEMENTED YET
     * protected void registerIndirectFunction() {
     * this.getWorkBook().addIndirectFormula(this);
     * } */


    /**
     * determine which formula in set of shared formula members is affected by cell br
     *
     * @param br cell
     * @return formula which references cell
     */
    fun getAffected(br: BiffRec): Formula? {
        val rc = IntArray(2)
        rc[0] = br.rowNumber
        rc[1] = br.colNumber.toInt()
        val ii = members!!.iterator()
        val isExcel2007 = this.workBook!!.isExcel2007
        while (ii.hasNext()) {
            val f = ii.next()
            val frc = IntArray(2)
            frc[0] = f.rowNumber
            frc[1] = f.colNumber.toInt()
            for (i in ptgcache!!.indices) {
                if (ptgcache!![i] is PtgRefN) {
                    val refrc = (ptgcache!![i] as PtgRefN).realRowCol
                    if (refrc[0] + frc[0] == rc[0] && adjustCol(refrc[1] + frc[1], isExcel2007) == rc[1])
                        return f
                } else {    // it's a PtgAreaN
                    val refrc = (ptgcache!![i] as PtgAreaN).realRowCol
                    refrc[0] += frc[0]
                    refrc[2] += frc[1]
                    refrc[1] = adjustCol(refrc[1] + frc[0], isExcel2007)
                    refrc[3] = adjustCol(refrc[3] + frc[0], isExcel2007)
                    if (refrc[0] <= rc[0] &&
                            refrc[1] <= rc[1] &&
                            refrc[2] >= rc[0] &&
                            refrc[3] >= rc[1])
                        return f
                }
            }
        }
        return null
    }

    /**
     * basic algorithm to adjust column dimensions when > MAXCOLS
     *
     * @param c
     * @param isExcel2007
     * @return
     */
    private fun adjustCol(c: Int, isExcel2007: Boolean): Int {
        var c = c
        if (c >= XLSConstants.MAXCOLS_BIFF8 && !isExcel2007)
            c -= XLSConstants.MAXCOLS_BIFF8
        return c
    }

    override fun close() {
        if (members != null)
            members!!.clear()
        members = null
        if (stack != null) {
            while (!stack!!.isEmpty()) {
                var p: GenericPtg? = stack!!.pop() as GenericPtg
                if (p is PtgRef)
                    p.close()
                else
                    p!!.close()
                p = null
            }
        }
        ptgcache = null
        host = null
        super.close()
    }

    protected fun finalize() {
        this.close()
    }

    companion object {
        private val serialVersionUID = -6147947203791941819L

        /**
         * Converts an expression stack that uses relative PTGs to
         * a standard stack for calculation
         */
        // NOTE: now these ptgs are not reference-tracked; see ExpressionParser and PtgRefN,PtgAreaN for reference-tracking these entities
        fun convertStack(`in`: Stack<*>, f: Formula): Stack<*> {
            val out = Stack()
            for (idx in `in`.indices) {
                var ptg = `in`[idx] as Ptg
                // convert the Ptg if necessary, otherwise clone it
                if (ptg is PtgRefN) {
                    ptg = ptg.convertToPtgRef(f)
                } else if (ptg is PtgAreaN) {
                    ptg = ptg.convertToPtgArea(f)
                } else {
                    ptg = ptg.clone() as Ptg
                    ptg.parentRec = f
                }

                out.add(ptg)
            }
            return out
        }
    }
}