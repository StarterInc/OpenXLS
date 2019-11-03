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
/*
   Ptg that indicates Unary minus, negates the operand on top of the stack
    
 * @see Ptg
 * @see Formula

    
*/
package io.starter.formats.XLS.formulas

import io.starter.toolkit.Logger

/**
 *
 */
class PtgUMinus : GenericPtg(), Ptg {

    override val isOperator: Boolean
        get() = true

    override val isPrimitiveOperator: Boolean
        get() = true

    override val isUnaryOperator: Boolean
        get() = true

    /**
     * return the human-readable String representation of
     */
    override val length: Int
        get() = Ptg.PTG_UMINUS_LENGTH

    override val string: String
        get() = "-"

    init {    // 20060504 KSC: Added to fill record bytes upon creation
        opcode = 0x13
        record = ByteArray(1)
        record[0] = 0x13
    }

    /**
     * Operator specific calculate method, this one returns a single value sent to it.
     */
    override fun calculatePtg(form: Array<Ptg>): Ptg? {
        if (form.size != 1) {
            Logger.logWarn("PtgMinus calculating formula failed, wrong number of values.")
            return PtgErr(PtgErr.ERROR_VALUE)
        }
        try {
            val p = form[0]
            var ret: Ptg? = null
            if (p is PtgInt) {
                var `val` = p.intVal
                `val` *= -1
                ret = PtgInt(`val`)
            } else {
                var `val` = p.doubleVal
                `val` *= -1.0
                ret = PtgNumber(`val`)
            }
            return ret
        } catch (e: Exception) {
            Logger.logWarn("PtgMinus calculating formula failed, could not negate operand " + form[0].toString() + " : " + e.toString())
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }

    override fun toString(): String {
        return "u-"
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 8448419489380791823L
    }

}
