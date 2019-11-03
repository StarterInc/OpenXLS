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

import io.starter.toolkit.ByteTools


/*
    Ptg that stores an integer value

    Offset  Name       Size     Contents
    ------------------------------------
    0       w           2       An unsigned integer value

 * @see Ptg
 * @see Formula


*/
class PtgInt : GenericPtg, Ptg {

    override val isOperand: Boolean
        get() = true

    override var intVal: Int = 0
        internal set

    /**
     * return the human-readable String representation of
     */
    override val string: String
        get() = intVal.toString()

    var `val`: Int
        get() = intVal
        set(i) {
            intVal = i
            this.updateRecord()
        }

    override val value: Any?
        get() = Integer.valueOf(intVal)

    val booleanVal: Boolean
        get() = intVal == 1

    override val length: Int
        get() = Ptg.PTG_INT_LENGTH

    constructor() {}

    /*
     * constructer to create ptgint's on the fly, from formulas
     */
    constructor(i: Int) {
        intVal = i
        this.updateRecord()
    }

    override fun init(b: ByteArray) {
        opcode = b[0]
        record = b
        this.populateVals()
    }

    // 0 to 65535 - outside of these bounds must be a PtgNumber
    private fun populateVals() {
        val b: Byte = 0
        val s = ByteTools.readInt(record[1], record[2], b, b)
        intVal = s
    }

    override fun updateRecord() {
        var tmp = ByteArray(1)
        tmp[0] = Ptg.PTG_INT
        val brow = ByteTools.shortToLEBytes(intVal.toShort())
        tmp = ByteTools.append(brow, tmp)
        record = tmp
    }

    override fun toString(): String {
        return this.`val`.toString()
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -2129624418815329359L
    }

}