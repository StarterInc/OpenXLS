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

/*
   Ptg that is a Concatenation Operand.
   Appends the top operand to the second-to-top Operand

 * @see Ptg
 * @see Formula


*/
class PtgConcat : GenericPtg(), Ptg {

    override val isOperator: Boolean
        get() = true

    override val isBinaryOperator: Boolean
        get() = true

    override val isPrimitiveOperator: Boolean
        get() = true    // 20060512 KSC: added

    /**
     * return the human-readable String representation of
     */
    override//        return "CONCAT(";	// 20060512 KSC: mod
    val string: String
        get() = "&"

    override//        return ")";
    val string2: String
        get() = ""

    override val length: Int
        get() = Ptg.PTG_CONCAT_LENGTH

    init {
        opcode = 0x8
        record = ByteArray(1)
        record[0] = 0x8
    }

    override fun toString(): String {    // KSC added
        return this.string
    }

    /**
     * Operator specific calculate method, this Concatenates two values
     */
    override fun calculatePtg(form: Array<Ptg>): Ptg? {
        try {
            val o = GenericPtg.getStringValuesFromPtgs(form)
            // there should always be only two ptg's in this, error if not.
            if (o == null || o.size != 2) {
                //if (o!=null)
                //	Logger.logWarn("calculating formula failed, wrong number of values in PtgConcat");
                return PtgErr(PtgErr.ERROR_VALUE)
            }
            if (!o[0].javaClass.isArray()) {
                val s = arrayOfNulls<String>(2)
                try {    // 20090216 KSC: try to convert numbers to ints when converting to string as otherwise all numbers come out as x.0
                    s[0] = (o[0] as Double).toInt().toString()
                } catch (e: Exception) {
                    s[0] = o[0].toString()
                }

                try {    // 20090216 KSC: try to convert numbers to ints when converting to string as otherwise all numbers come out as x.0
                    s[1] = (o[1] as Double).toInt().toString()
                } catch (e: Exception) {
                    s[1] = o[1].toString()
                }

                val returnVal = s[0] + s[1]
                return PtgStr(returnVal)
            } else {
                return null
            }
        } catch (e: Exception) {    // handle error ala Excel
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 6671404163121438253L
    }
}