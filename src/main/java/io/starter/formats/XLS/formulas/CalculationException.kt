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

/**
 * Indicates an error occurred during formula calculation.
 */
open class CalculationException
/**
 * Creates a new CaluculationException.
 *
 * @param error the error code. must be one of the defined error constants.
 */
(
        /**
         * The error code for this error.
         */
        private val error: Byte) : Exception() {

    /**
     * Gets the BIFF8 error code for this error.
     */
    val errorCode: Byte
        get() = if (error == CIR_ERR) VALUE else error

    /**
     * Gets the string name of this error.
     */
    open val name: String?
        get() {
            when (error) {
                NULL -> return "#NULL!"
                DIV0 -> return "#DIV/0!"
                VALUE -> return "#VALUE!"
                REF -> return "#REF!"
                NAME -> return "#NAME?"
                NUM -> return "#NUM!"
                NA -> return "#N/A"
                CIR_ERR -> return "#CIR_ERR!"
                else -> return null
            }
        }

    /**
     * Gets a human-readable message describing this error.
     */
    override fun getMessage(): String {
        when (error) {
            NULL -> return "a range intersection returned no cells"
            DIV0 -> return "attempted to divide by zero"
            VALUE -> return "operand type mismatch"
            REF -> return "reference to a cell that doesn't exist"
            NAME -> return "reference to an unknown function or defined name"
            NUM -> return "number storage overflow"
            NA -> return "lookup returned no value for the given criteria"
            CIR_ERR -> return "circular reference error"
            else -> return "unknown error occurred"
        }
    }

    override fun toString(): String? {
        return name
    }

    companion object {
        private val serialVersionUID = 2028428287133817627L
        internal var errorStrings = arrayOf(arrayOf("#DIV/0!", "7"), arrayOf("#N/A", "42"), arrayOf("#NAME?", "29"), arrayOf("#NULL!", "0"), arrayOf("#NUM!", "36"), arrayOf("#REF!", "23"), arrayOf("#VALUE!", "15"), arrayOf("#CIR_ERR!", "15")  // false error with code for value, output for circular ref exceptions
        )

        /**
         * Excel #NULL! error.
         * Indicates that a range intersection returned no cells.
         */
        val NULL = 0x00.toByte()

        /**
         * Excel #DIV/0! error.
         * Indicates that the formula attempted to divide by zero.
         */
        val DIV0 = 0x07.toByte()

        /**
         * Excel #VALUE! error.
         * Indicates that there was an operand type mismatch.
         */
        val VALUE = 0x0F.toByte()

        /**
         * Excel #REF! error.
         * Indicates that a reference was made to a cell that doesn't exist.
         */
        val REF = 0x17.toByte()

        /**
         * Excel #NAME? error.
         * Indicates an unknown string was encountered in the formula.
         */
        val NAME = 0x1D.toByte()

        /**
         * Excel #NUM! error.
         * Indicates that a calculation result overflowed the number storage.
         */
        val NUM = 0x24.toByte()

        /**
         * Excel #N/A error.
         * Indicates that a lookup (e.g. VLOOKUP) returned no results.
         */
        val NA = 0x2A.toByte()

        /**
         * Custom circular exception error, internally stores as a #VALUE
         */
        val CIR_ERR = 0xFF.toByte()

        /**
         * static version, takes String error code and returns the correct error code
         *
         * @param error String
         * @return
         */
        fun getErrorCode(error: String?): Byte {
            if (error == null) return 0    // unknown
            for (i in errorStrings.indices) {
                if (error == errorStrings[i][0]) {
                    return Byte(errorStrings[i][1])
                }
            }
            return 0
        }
    }
}
