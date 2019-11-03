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
   Indicate the placing of operatands in parenthesis.

   Does not affect calculation.

   =1+(2)

   ptgInt 1
   ptgInt 2
   ptgParen
   ptgAdd

   =(1+2)

   ptgInt 1
   ptgInt 2
   ptgAdd
   ptgParen


 * @see Ptg
 * @see Formula

  .
*/
class PtgParen : GenericPtg(), Ptg {

    override val isControl: Boolean
        get() = true

    /**
     * return the human-readable String representation of
     */
    override// TODO: add logic to return proper paren 12/02 -jm
    val string: String
        get() = "("

    /**
     * return the human-readable String representation of
     * the "closing" portion of this Ptg
     * such as a closing parenthesis.
     */
    override val string2: String
        get() = ")"

    override val length: Int
        get() = Ptg.PTG_PAREN_LENGTH


    /**
     * Pass in the last 3 ptgs to evaluate
     * where to place the String parens.
     */
    override fun evaluate(b: Array<Any>): Any? {
        return null
    }

    // KSC: added
    //default constructor
    init {
        opcode = 0x15
        record = ByteArray(1)
        record[0] = opcode
    }

    override fun toString(): String {
        return ")"
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -8081397558698615537L
    }
}