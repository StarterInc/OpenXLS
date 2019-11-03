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


/*
   Ptg that is a multiplier operand

 * @see Ptg
 * @see Formula


*/
class PtgMlt : GenericPtg(), Ptg {

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
        get() = "*"

    override val length: Int
        get() = Ptg.PTG_MLT_LENGTH

    override fun toString(): String {
        return this.string
    }

    init {
        opcode = 0x5
        record = ByteArray(1)
        record[0] = 0x5
    }

    /*  Operator specific calculate method, this one multiplies two values.

     */
    override fun calculatePtg(form: Array<Ptg>): Ptg? {
        // handle ref errs
        if (form[0] is PtgErr || form[0] is PtgRefErr || form[0] is PtgRefErr3d || form[0] is PtgAreaErr3d)
            return form[0]
        if (form[1] is PtgErr || form[1] is PtgRefErr || form[1] is PtgRefErr3d || form[1] is PtgAreaErr3d)
            return form[1]

        try {
            // 20090202 KSC: Handle array formulas
            val o = GenericPtg.getValuesFromPtgs(form) ?: return PtgErr(PtgErr.ERROR_VALUE)
// some error in value(s)
            if (!o[0].javaClass.isArray()) {
                //double[] dub = super.getValuesFromPtgs(form);
                // there should always be only two ptg's in this, error if not.
                if (o.size != 2) {
                    Logger.logWarn("calculating formula failed, wrong number of values in PtgMlt")
                    return PtgErr(PtgErr.ERROR_VALUE)    // 20081203 KSC: handle error's ala Excel return null;
                }
                var o0 = 0.0
                var o1 = 0.0
                try {
                    o0 = GenericPtg.getDoubleValue(o[0], this.parentRec)
                    o1 = GenericPtg.getDoubleValue(o[1], this.parentRec)
                } catch (e: NumberFormatException) {
                    return PtgErr(PtgErr.ERROR_VALUE)
                }

                val returnVal = o0 * o1
                // create a container ptg for these.
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
                        var o0 = 0.0
                        var o1 = 0.0
                        try {
                            o0 = GenericPtg.getDoubleValue(firstOp, this.parentRec)
                            o1 = GenericPtg.getDoubleValue(secondOp, this.parentRec)
                        } catch (e: NumberFormatException) {
                            retArry = "$retArry#VALUE!,"
                            continue
                        }

                        val retVal = o0 * o1
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
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -2670754297349356254L
    }

}