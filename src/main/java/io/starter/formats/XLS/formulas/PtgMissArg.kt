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
   Ptg that is an missing operand

 * @see Ptg
 * @see Formula


*/

class PtgMissArg : GenericPtg(), Ptg {

    override val isOperand: Boolean
        get() = true

    override val isOperator: Boolean
        get() = false

    override val isBinaryOperator: Boolean
        get() = false

    override val isPrimitiveOperator: Boolean
        get() = false

    /**
     * return the human-readable String representation of
     */
    override val string: String
        get() = ""

    override val length: Int
        get() = 1

    override val value: Any?
        get() = null

    init {
        this.init(byteArrayOf(22))
    }

    override fun toString(): String {
        return ""
    }

    companion object {

        private val serialVersionUID = 8995314621921283625L
    }
}