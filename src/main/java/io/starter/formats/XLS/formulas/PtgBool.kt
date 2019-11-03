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

import io.starter.formats.XLS.ExpressionParser

/*
    A parse thing that represents a boolean value.  This is made up of two bytes,
    the PtgID (0x1D) and a byte representing the boolean value (0 or 1);
*/

class PtgBool : GenericPtg, Ptg {

    override val isOperator: Boolean
        get() = false

    override val isOperand: Boolean
        get() = true

    var booleanValue = false
        internal set

    /**
     * return the human-readable String representation of the ptg
     */
    override val string: String
        get() = booleanValue.toString()

    override val value: Any?
        get() = java.lang.Boolean.valueOf(booleanValue)

    override val length: Int
        get() = Ptg.PTG_BOOL_LENGTH

    override fun toString(): String {
        return booleanValue.toString()
    }

    fun setVal(boo: Boolean) {
        booleanValue = boo
        this.updateRecord()
    }

    override fun init(rec: ByteArray) {
        this.record = rec
        opcode = rec[0]
        booleanValue = rec[1].toInt() != 0
    }

    constructor() {
        record = ByteArray(2)
        opcode = ExpressionParser.ptgBool.toByte()
        record[0] = opcode
    }


    constructor(b: Boolean) {
        opcode = ExpressionParser.ptgBool.toByte()
        booleanValue = b
        this.updateRecord()
    }

    override fun updateRecord() {
        record = ByteArray(2)
        record[0] = opcode
        if (booleanValue) {
            record[1] = 1
        } else {
            record[1] = 0
        }
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -7270271326251770439L
    }

}