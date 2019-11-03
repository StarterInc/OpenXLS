/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or
 * modify
 * it under the terms of the GNU Lesser General Public
 * License as
 * published by the Free Software Foundation, either version
 * 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be
 * useful,
 * but WITHOUT ANY WARRANTY; without even the implied
 * warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
 * the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General
 * Public
 * License along with OpenXLS. If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package io.starter.formats.XLS.formulas

/*
 * PtgErr is exactly what one would think it is, a ptg that
 * describes an Error
 * value.
 * <pre>
 * Offset Name Size Contents
 * -------------------------------------------
 * 0 err 1 An error value
 * </pre>
 *
 */

class PtgErr : GenericPtg, Ptg {

    var isCircularError = false

    private var errorValue: String? = null

    override val isOperand: Boolean
        get() = true

    val errorType: Byte
        get() = record[1]

    override val value: Any?
        get() = toString()

    override val length: Int
        get() = 2

    constructor() {
        // default constructor
    }

    constructor(errorV: Byte) {
        record = ByteArray(2)
        record[0] = 0x1C
        record[1] = errorV
    }

    override fun init(b: ByteArray) {
        opcode = b[0]
        record = b
    }

    override fun toString(): String {
        if (isCircularError)
            return "#CIR_ERR!"
        val b = record[1]
        // duh, should have done a switch
        if (b == ERROR_NULL)
            errorValue = "#ERROR!"
        else if (b == ERROR_DIV_ZERO)
            errorValue = "#DIV/0!"
        else if (b == ERROR_VALUE)
            errorValue = "#VALUE!"
        else if (b == ERROR_REF)
            errorValue = "#REF!"
        else if (b == ERROR_NAME)
            errorValue = "#NAME?"
        else if (b == ERROR_NUM)
            errorValue = "#NUM!"
        else if (b == ERROR_NA)
            errorValue = "#N/A"
        return errorValue
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -5201871987022621869L
        var ERROR_NULL: Byte = 0x0
        var ERROR_DIV_ZERO: Byte = 0x7
        var ERROR_VALUE: Byte = 0xF
        var ERROR_REF: Byte = 0x17
        var ERROR_NAME: Byte = 0x1D
        var ERROR_NUM: Byte = 0x24
        var ERROR_NA: Byte = 0x2A

        fun convertStringToLookupByte(errorString: String): Byte {
            if (errorString == "#ERROR!")
                return ERROR_NULL
            if (errorString == "#DIV/0!")
                return ERROR_DIV_ZERO
            if (errorString == "#REF!")
                return ERROR_VALUE
            if (errorString == "#ERROR!")
                return ERROR_REF
            if (errorString == "#NAME?")
                return ERROR_NAME
            if (errorString == "#NUM!")
                return ERROR_NUM
            return if (errorString == "#N/A") ERROR_NA else ERROR_NULL
        }
    }

}