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

import java.util.ArrayList


/*
 * Computes the Union of the two top operands. Supposedly this is
 * "Microsoft Excel's comma operator" Brought to you by the same people that
 * created those space saving RK's
 *
 *
 * "http://www.extentech.com">Extentech Inc.</a>
 *
 *
 *
 * @see Ptg
 *
 * @see Formula
 */
class PtgUnion : GenericPtg(), Ptg {

    override val isOperator: Boolean
        get() = true

    override val isBinaryOperator: Boolean
        get() = true

    override val isPrimitiveOperator: Boolean
        get() = true

    /**
     * return the human-readable String representation of
     */
    override val string: String
        get() = ","

    override val length: Int
        get() = Ptg.PTG_UNION_LENGTH

    init {
        opcode = 0x10
        record = ByteArray(1)
        record[0] = 0x10
    }

    /**
     * Union = All items from A and B, shared and not shared. Range list
     * operator, represented by the system's list separator sign (for example
     * comma sign). Treats two ranges as one operator (A1:B2,B2:C3) B1:B3
     * ==>B1:B2, B2:B3 (A1:B2,B2:C3) ==>A1:B2, B2:C3
     */
    // just add together? seems that is the case
    override fun calculatePtg(form: Array<Ptg>): Ptg? {
        if (form.size != 2)
            return PtgErr(PtgErr.ERROR_VALUE)

        try {
            var sourceSheet: String? = null
            try {
                if (this.parentRec!!.sheet != null)
                // could be a Name rec ...
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
                } else if (p is PtgRef3d) {
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
                        if (!loc.startsWith("#")) { // skip errors -- TODO: IS THIS CORRECT?
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
                    }
                } else if (p is PtgErr || p is PtgRefErr || p is PtgAreaErr3d) {
                    // DO WHAT???
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
            // now have components for both operands - SUM and package into 1 ptg for return
            // For performance reasons, instantiate a PtgMystery as a lightweight GenericPtg which holds intermediary values in it's vars
            val retp = PtgMystery()
            val retptgs = ArrayList()
            for (k in first!!.indices)
                retptgs.add(first[k])
            for (k in last!!.indices)
                retptgs.add(last[k])
            val ptgs = arrayOfNulls<Ptg>(retptgs.size)
            retptgs.toTypedArray()
            retp.setVars(ptgs)
            return retp
        } catch (e: NumberFormatException) {
            return PtgErr(PtgErr.ERROR_VALUE)
        } catch (e: Exception) {    // handle error ala Excel
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 8333035819099274707L
    }
}