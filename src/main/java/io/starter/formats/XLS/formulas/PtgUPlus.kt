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
   Ptg that indicates Unary plus, has no effect on operand.  Very useful PTG...lol
    
 * @see Ptg
 * @see Formula

    
*/
package io.starter.formats.XLS.formulas

import io.starter.toolkit.Logger

/**
 *
 */
class PtgUPlus : GenericPtg(), Ptg {

    override val isOperator: Boolean
        get() = true

    override val isPrimitiveOperator: Boolean
        get() = true

    override val isUnaryOperator: Boolean
        get() = true

    /**
     * return the human-readable String representation of
     */
    override val string: String
        get() = "+"

    override val length: Int
        get() = Ptg.PTG_UPLUS_LENGTH

    init {    // 20060504 KSC: Added to fill record bytes upon creation
        opcode = 0x12
        record = ByteArray(1)
        record[0] = 0x12
    }

    /*  Operator specific calculate method, this one returns a single value sent to it.

     */
    override fun calculatePtg(form: Array<Ptg>): Ptg? {
        // there should always be only one ptg in this, error if not.
        if (form.size != 1) {
            Logger.logWarn("calculating formula failed, wrong number of values in PtgUPlus")
            return null
        }
        return form[0]
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -3514760881731524419L
    }

}
