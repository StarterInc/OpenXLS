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
   Equals operand

   Evaluates to true if the top two operands are equal, otherwise FALSE


 * @see Ptg
 * @see Formula


*/
class PtgEQ : GenericPtg(), Ptg {

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
        get() = "="

    override val length: Int
        get() = Ptg.PTG_EQ_LENGTH

    init {
        opcode = 0xB
        record = ByteArray(1)
        record[0] = 0xB
    }

    override fun toString(): String {
        return this.string
    }

    /*  Operator specific calculate method, this one determines if the second-to-top
       operand is equal to the top operand;  Returns a PtgBool

   */
    override fun calculatePtg(form: Array<Ptg>): Ptg? {
        var res = false
        // there should always be only two ptg's in this, error if not.
        if (form.size != 2) {
            Logger.logInfo("calculating formula, wrong number of values in PtgEQ")
            return PtgErr(PtgErr.ERROR_VALUE)    // 20081203 KSC: handle error's ala Excel
        }
        // check for null referenced values, a null reference is equal to the string "";
        val o = GenericPtg.getValuesFromPtgs(form) ?: return PtgErr(PtgErr.ERROR_VALUE)
// some error in value(s)
        if (o[1].javaClass.isArray() && !o[0].javaClass.isArray()) {
            val tmp = o[0]
            o[0] = o[1]
            o[1] = tmp
        }
        if (!o[0].javaClass.isArray()) {
            if (o.size != 2)
                return PtgErr(PtgErr.ERROR_VALUE)    // 20081203 KSC: handle error's ala Excel return null;
            // blank handling:
            // determine if any of the operands are double - if true,
            // then blank comparisons will be treated as 0's
            var isDouble = false
            run {
                var i = 0
                while (i < 2 && !isDouble) {
                    //if (!form[i].isBlank())
                    isDouble = o[i] is Double
                    i++
                }
            }
            for (i in 0..1) {
                //if (form[i].isBlank()) {
                if (o[i] != null && o[i].toString() == "") {
                    if (isDouble)
                        o[i] = 0.0
                    else
                        o[i] = "" // in this case, empty cells are handled as blank, not zero
                }
            }
            if (o[0] === o[1])
                res = true
            else if (o[0] == null || o[1] == null)
                res = false
            else if (o[0] is Double && o[1] is Double)
                res = Math.abs((o[0] as Double).toDouble() - (o[1] as Double).toDouble()) < doublePrecision    // compare equality to certain precision
            else if (o[0].toString().equals(o[1].toString(), ignoreCase = true))
                res = true
            // handle empty cell references vs string case, 0.0 does not match
        } else {    // handle array formulas
            var retArry = ""
            val nArrays = java.lang.reflect.Array.getLength(o)
            if (nArrays != 2) return PtgErr(PtgErr.ERROR_VALUE)
            val nVals = java.lang.reflect.Array.getLength(o[0])    // use first array element to determine length of values as subsequent vals might not be arrays
            if (nVals == 0) {
                retArry = "{false}"
                val pa = PtgArray()
                pa.setVal(retArry)
                return pa
            }
            var i = 0
            while (i < nArrays - 1) {
                res = false
                var secondOp: Any? = null
                val comparitorIsArray = o[i + 1].javaClass.isArray()
                if (!comparitorIsArray) secondOp = o[i + 1]
                for (j in 0 until nVals) {
                    val firstOp = Array.get(o[i], j)    // first array index j
                    if (comparitorIsArray)
                        secondOp = Array.get(o[i + 1], j)    // second array index j
                    var fd = 0.0
                    var sd = 0.0
                    try {
                        fd = Double(firstOp.toString())
                        sd = Double(secondOp!!.toString())
                        res = Math.abs(fd - sd) <= doublePrecision    // compare to certain precision instead of equality

                    } catch (e: Exception) {
                        //if (firstOp instanceof Double && secondOp instanceof Double)
                        //res= (Math.abs((((Double)firstOp).doubleValue())-((Double)secondOp).doubleValue()))<=doublePrecision;	// compare to certain precision instead of equality
                        //else
                        res = firstOp.toString().equals(secondOp!!.toString(), ignoreCase = true)
                    }

                    retArry = "$retArry$res,"
                }
                i += 2
            }
            retArry = "{" + retArry.substring(0, retArry.length - 1) + "}"
            val pa = PtgArray()
            pa.setVal(retArry)
            return pa
        }
        return PtgBool(res)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 5446048862531696036L
    }

}