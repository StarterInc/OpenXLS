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
  Ptg that is a not equal to operand

   Evaluates to TRUE if the two top operands are not equal,
   otherwise evaluates as FALSE;


 * @see Ptg
 * @see Formula


*/
class PtgNE : GenericPtg(), Ptg {

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
        get() = "<>"

    override val length: Int
        get() = Ptg.PTG_NE_LENGTH

    init {
        opcode = 0xE
        record = ByteArray(1)
        record[0] = 0xE
    }

    /*  Operator specific calculate method, this one determines if the second-to-top
       operand is less than the top operand;  Returns a PtgBool

   */
    override fun calculatePtg(form: Array<Ptg>): Ptg? {
        try {
            // 20090202 KSC: Handle array formulas
            val o = GenericPtg.getValuesFromPtgs(form)
            var res: Boolean
            if (!o!![0].javaClass.isArray()) {
                //double[] dub = super.getValuesFromPtgs(form);
                // there should always be only two ptg's in this, error if not.
                if (o == null || o.size != 2) {
                    Logger.logWarn("calculating formula failed, wrong number of values in PtgNE")
                    return null
                }
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
                if (o[0] is Double && o[1] is Double)
                    res = Math.abs((o[0] as Double).toDouble() - (o[1] as Double).toDouble()) > doublePrecision    // compare equality to certain precision
                else
                    res = !o[0].toString().equals(o[1].toString(), ignoreCase = true)
                return PtgBool(res)
            } else {    // handle array fomulas
                var retArry = ""
                val nArrays = java.lang.reflect.Array.getLength(o)
                if (nArrays != 2) return PtgErr(PtgErr.ERROR_VALUE)
                val nVals = java.lang.reflect.Array.getLength(o[0])    // use first array element to determine length of values as subsequent vals might not be arrays
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

                        if (firstOp is Double && secondOp is Double)
                            res = Math.abs(firstOp.toDouble() - secondOp.toDouble()) > doublePrecision    // compare to certain precision instead of equality
                        else
                            res = firstOp.toString().equals(secondOp!!.toString(), ignoreCase = true)
                        retArry = "$retArry$res,"
                    }
                    i += 2
                }
                retArry = "{" + retArry.substring(0, retArry.length - 1) + "}"
                val pa = PtgArray()
                pa.setVal(retArry)
                return pa
            }
            /*}catch(NumberFormatException e){	// shouldn't get here!! see new code above
    		String[] s = getStringValuesFromPtgs(form);
    		if (s[0].equalsIgnoreCase(s[1]))return new PtgBool(false);
    		return new PtgBool(true);
    	}*/
        } catch (e: Exception) {    // 20090212 KSC: handle error ala Excel
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 6901661166166179786L
    }

}