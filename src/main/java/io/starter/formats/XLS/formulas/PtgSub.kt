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

import io.starter.toolkit.Logger

import java.lang.reflect.Array


/**
 * Ptg that indicates substitution (ie minus)
 *
 * @see Ptg
 *
 * @see Formula
 */
class PtgSub : GenericPtg(), Ptg {

    override val isOperator: Boolean
        get() = true

    override val isPrimitiveOperator: Boolean
        get() = true

    override val isBinaryOperator: Boolean
        get() = true

    /**
     * return the human-readable String representation of
     */
    override val string: String
        get() = "-"

    override val length: Int
        get() = Ptg.PTG_SUB_LENGTH

    init {
        opcode = 0x4
        record = ByteArray(1)
        record[0] = 0x4
    }

    override fun toString(): String {
        return string
    }

    /**
     * Operator specific calculate method, this one subtracts one value from another
     */
    override fun calculatePtg(form: Array<Ptg>): Ptg? {
        try {
            val o = GenericPtg.getValuesFromPtgs(form) ?: return PtgErr(PtgErr.ERROR_VALUE)
// some error in value(s)
            if (!o[0].javaClass.isArray()) {
                if (o.size != 2) {
                    Logger.logWarn("calculating formula failed, wrong number of values in PtgSub")
                    return PtgErr(PtgErr.ERROR_VALUE)    // 20081203 KSC: handle error's ala Excel return null;
                }
                // blank handling:
                if (form[0].isBlank) o[0] = 0
                if (form[1].isBlank) o[1] = 0
                // the following should only return #VALUE! if ???
                if (!(o[0] is Double && o[1] is Double)) {
                    return if (this.parentRec == null) {
                        PtgErr(PtgErr.ERROR_VALUE)
                    } else if (this.parentRec!!.sheet!!.window2!!.showZeroValues) {
                        PtgInt(0)
                    } else {
                        PtgStr("")
                    }
                }
                val returnVal = (o[0] as Double).toDouble() - (o[1] as Double).toDouble()
                return PtgNumber(returnVal)
            } else {    // handle array fomulas
                var retArry = ""
                val nArrays = java.lang.reflect.Array.getLength(o)
                if (nArrays != 2) return PtgErr(PtgErr.ERROR_VALUE)
                val nVals = java.lang.reflect.Array.getLength(o[0])    // use first array element to determine length of values as subsequent vals might not be arrays
                var i = 0
                while (i < nArrays - 1) {
                    var secondOp: Any? = null
                    val comparitorIsArray = o[i + 1].javaClass.isArray()
                    if (!comparitorIsArray) secondOp = o[i + 1]
                    for (j in 0 until nVals) {
                        val firstOp = Array.get(o[i], j)    // first array index j
                        if (comparitorIsArray)
                            secondOp = Array.get(o[i + 1], j)    // second array index j
                        if (!(firstOp is Double && secondOp is Double)) {
                            return if (this.parentRec == null) {
                                PtgErr(PtgErr.ERROR_VALUE)
                            } else if (this.parentRec!!.sheet!!.window2!!.showZeroValues) {
                                PtgInt(0)
                            } else {
                                PtgStr("")
                            }
                        }    // 20081203 KSC: handle error's ala Excel
                        val retVal = firstOp.toDouble() - secondOp.toDouble()
                        retArry = "$retArry$retVal,"
                    }
                    i += 2
                }
                retArry = "{" + retArry.substring(0, retArry.length - 1) + "}"
                val pa = PtgArray()
                pa.setVal(retArry)
                return pa
            }
        } catch (e: NumberFormatException) {
            return PtgErr(PtgErr.ERROR_VALUE)
        } catch (e: Exception) {    // 20081125 KSC: handle error ala Excel
            // Logger.logErr("PtgSub failed:" + e);
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -3252464873846778499L
    }
}