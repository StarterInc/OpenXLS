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


/*
   Ptg that is a exponent operand
   Raises the second-to-top operand to the power of the
   top operand

 * @see Ptg
 * @see Formula


*/
class PtgPower : GenericPtg(), Ptg {

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
        get() = "^"

    override val length: Int
        get() = Ptg.PTG_POWER_LENGTH

    init {
        opcode = 0x7
        record = ByteArray(1)
        record[0] = 0x7
    }

    /*  Operator specific calculate method, this one raises the second-to-top
    operand to the power of the top operand

*/
    override fun calculatePtg(form: Array<Ptg>): Ptg? {
        try {
            // 20090202 KSC: Handle array formulas
            val o = GenericPtg.getValuesFromPtgs(form)
            if (!o!![0].javaClass.isArray()) {
                //double[] dub = super.getValuesFromPtgs(form);
                // there should always be only two ptg's in this, error if not.
                if (o == null || o.size != 2) {
                    Logger.logWarn("calculating formula failed, wrong number of values in PtgPower")
                    return null
                }
                //double returnVal = Math.pow(dub[0].doubleValue(), dub[1].doubleValue());
                val returnVal = Math.pow((o[0] as Double).toDouble(), (o[1] as Double).toDouble())
                // create a container ptg for these.
                return PtgNumber(returnVal)
            }
            // TODO: FINISH ARRAY FORMULAS
            return PtgErr(PtgErr.ERROR_VALUE)
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 4675566993519011450L
    }
}