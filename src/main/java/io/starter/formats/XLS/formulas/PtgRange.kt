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

import io.starter.OpenXLS.ExcelTools

import java.util.ArrayList


/*
   Computes the minimal bounding rectangle of the top two operands.
   This is excel's ":" colon operator.


 * @see Ptg
 * @see Formula


*/
class PtgRange : GenericPtg(), Ptg {

    override val isOperator: Boolean
        get() = true

    override val isBinaryOperator: Boolean
        get() = true

    override val isPrimitiveOperator: Boolean
        get() = true    // 20091019 KSC

    /**
     * return the human-readable String representation of
     */
    override val string: String
        get() = ":"


    override val length: Int
        get() = Ptg.PTG_RANGE_LENGTH
    /*? public boolean getIsOperand(){return true;}
    public boolean getIsControl(){return true;}
    */

    init {
        opcode = 0x11
        record = ByteArray(1)
        record[0] = 0x11
    }

    /**
     * The RANGE operator (:)
     * Returns the minimal rectangular range that contains both parameters.
     * A1:B2:B2:C3 ==>A1:C3
     *
     *
     * NOTE: assumption is NO 3d refs i.e. all on same sheet *******
     */
    override fun calculatePtg(form: Array<Ptg>): Ptg? {
        if (form.size != 2)
            return PtgErr(PtgErr.ERROR_VALUE)

        try {
            var sheet: String? = null
            var sourceSheet: String? = null
            try {
                sourceSheet = this.parentRec!!.sheet!!.sheetName
            } catch (ne: NullPointerException) {
            }

            var first: ArrayList<*>? = null
            var last: ArrayList<*>? = null
            for (i in 0..1) {
                val p = form[i]
                val a = ArrayList()
                if (p is PtgArea) {
                    a.add(p)
                } else if (p is PtgRef) {
                    a.add(p)
                } else if (p is PtgName) {
                    val pc = p.components
                    for (j in pc.indices)
                        a.add(pc[j])
                } else if (p is PtgStr) {
                    val comps = p.toString().split(",".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
                    for (j in comps.indices) {
                        if (comps[j].indexOf(":") == -1) {
                            if (comps[j] != "#REF!" && comps[j] != "#NULL!") {
                                val pr = PtgRef3d(false)
                                pr.parentRec = this.parentRec
                                pr.location = comps[j]
                                a.add(pr)
                            } else {
                                val pr = PtgRefErr3d()
                                pr.parentRec = this.parentRec
                                a.add(pr)
                            }
                        } else {
                            val pa = PtgArea3d(false)
                            pa.parentRec = this.parentRec
                            pa.location = comps[j]
                            val pcs = pa.components
                            if (pcs != null) {
                                for (k in pcs.indices) {
                                    (pcs[k] as PtgRef).sheetName = pa.sheetName
                                    a.add(pcs[k])
                                }
                            }
                        }
                    }
                } else if (p is PtgArray) {
                    // parse array components and create refs
                    val pc = p.components
                    for (j in pc.indices) {
                        val loc = pc[j].toString()
                        if (loc.indexOf(":") == -1) {
                            if (loc.indexOf("!") == -1) {
                                val pr = PtgRef()
                                pr.useReferenceTracker = false
                                pr.parentRec = this.parentRec
                                pr.location = loc
                                a.add(pr)
                            } else {
                                val pr = PtgRef3d(false)
                                pr.parentRec = this.parentRec
                                pr.location = loc
                                a.add(pr)
                            }
                        } else {
                            val pa = PtgArea3d(false)
                            pa.parentRec = this.parentRec
                            pa.location = loc
                            a.add(pa)
                        }
                    }
                } else if (p is PtgErr || p is PtgRefErr || p is PtgAreaErr3d) {
                    // DO WHAT???
                    // ignore
                } else {        // if an intermediary value returned from PtgRange, PtgUnion or PtgIsect, will be a GenericPtg which holds intermediary values in its vars array
                    val pc = (p as GenericPtg).vars
                    for (j in pc!!.indices) {
                        if ((pc[j] is PtgArea) and (pc[j] !is PtgAreaErr3d)) {
                            val pa = pc[j].components
                            for (k in pa.indices)
                                a.add(pa[k])
                        } else
                            a.add(pc[j])
                    }
                }

                if (first == null)
                    first = a
                else
                    last = a
            }
            // now have components for both operands
            // range op returns the range that encompasses all referenced ptgs
            val rng = intArrayOf(java.lang.Short.MAX_VALUE.toInt(), java.lang.Short.MAX_VALUE.toInt(), 0, 0)
            for (k in first!!.indices) {
                val pr = first[k] as PtgRef
                if (sheet == null) sheet = pr.sheetName    // TODO: 3d ranges??????
                val rc = pr.intLocation
                if (rc!!.size > 2) { // it's a range
                    val numrows = rc[2] - rc[0] + 1
                    val numcols = rc[3] - rc[1] + 1
                    var numcells = numrows * numcols
                    if (numcells < 0)
                        numcells *= -1 // handle swapped cells ie: "B1:A1"
                    var rowctr = rc[0]
                    var cellctr = rc[1] - 1
                    for (i in 0 until numcells) {
                        if (cellctr == rc[3]) {// if its the end of the row,increment row.
                            cellctr = rc[1] - 1
                            rowctr++
                        }
                        ++cellctr
                        val addr = intArrayOf(rowctr, cellctr)
                        adjustRange(addr, rng)
                    }
                } else {
                    adjustRange(rc, rng)
                }
            }
            for (k in last!!.indices) {
                val pr = last[k] as PtgRef
                if (sheet == null) sheet = pr.sheetName    // TODO: 3d ranges??????
                val rc = pr.intLocation
                if (rc!!.size > 2) { // it's a range
                    if (rc.size > 2) { // it's a range
                        val numrows = rc[2] - rc[0] + 1
                        val numcols = rc[3] - rc[1] + 1
                        var numcells = numrows * numcols
                        if (numcells < 0)
                            numcells *= -1 // handle swapped cells ie: "B1:A1"
                        var rowctr = rc[0]
                        var cellctr = rc[1] - 1
                        for (i in 0 until numcells) {
                            if (cellctr == rc[3]) {// if its the end of the row,increment row.
                                cellctr = rc[1] - 1
                                rowctr++
                            }
                            ++cellctr
                            val addr = intArrayOf(rowctr, cellctr)
                            adjustRange(addr, rng)
                        }
                    }
                } else {
                    adjustRange(rc, rng)
                }
            }
            // For performance reasons, instantiate a PtgMystery as a lightweight GenericPtg which holds intermediary values in it's vars
            val retp = PtgMystery()
            val pa = PtgArea3d(false)
            pa.parentRec = this.parentRec
            // TODO: 3d ranges????
            pa.sheetName = sheet
            pa.setLocation(rng)
            retp.setVars(arrayOf<Ptg>(pa))
            return retp
        } catch (e: NumberFormatException) {
            return PtgErr(PtgErr.ERROR_VALUE)
        } catch (e: Exception) {    // handle error ala Excel
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }


    private fun adjustRange(rc: IntArray, rng: IntArray) {
        if (ExcelTools.isBeforeRange(rc, rng)) {
            rng[0] = rc[0]
            rng[1] = rc[1]
        }
        if (ExcelTools.isAfterRange(rc, rng)) {
            rng[2] = rc[0]
            rng[3] = rc[1]
        }
    }

    companion object {

        private val serialVersionUID = 7181427387507157013L
    }

}