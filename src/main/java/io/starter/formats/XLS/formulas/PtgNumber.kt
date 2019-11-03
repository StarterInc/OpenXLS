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
    Ptg that stores an IEEE value

    Offset  Name       Size     Contents
    ------------------------------------
    0       num          8      An IEEE floating point nubmer

 * @see Ptg
 * @see Formula


*/
class PtgNumber : GenericPtg, Ptg {

    override val isOperand: Boolean
        get() = true

    /**
     * override of GenericPtg.getDoubleVal();
     */
    override var doubleVal: Double = 0.toDouble()
        internal set
    internal var percentage = false    // 20081208 KSC: so can handle percentage values in String formulas

    /**
     * return the human-readable String representation of
     */
    override val string: String
        get() = if (!percentage) doubleVal.toString() else (doubleVal * 100).toString() + "%"

    override val value: Any?
        get() = doubleVal

    var `val`: Double
        get() = doubleVal
        set(i) {
            doubleVal = i
            this.updateRecord()
        }

    override val length: Int
        get() = Ptg.PTG_NUM_LENGTH

    constructor() {
        opcode = 0x1F
        doubleVal = 0.0
        this.updateRecord()
    }

    override fun init(b: ByteArray) {
        opcode = b[0]
        record = b
        this.populateVals()
    }

    /**
     * Constructer to create these on the fly, this is needed
     * for value storage in calculations of formulas.
     */
    constructor(d: Double) {
        opcode = 0x1F
        doubleVal = d
        this.updateRecord()
    }


    private fun populateVals() {
        val barr = ByteArray(8)
        System.arraycopy(record, 1, barr, 0, 8)
        doubleVal = ByteTools.eightBytetoLEDouble(barr)
    }

    // 20081208 KSC: handle percentage values
    fun setVal(s: String) {
        var s = s
        s = s.trim { it <= ' ' }
        if (s.indexOf("%") == s.length - 1) {
            percentage = true
            s = s.substring(0, s.indexOf("%"))
            doubleVal = Double(s) / 100
        } else
            doubleVal = Double(s)
    }

    override fun updateRecord() {
        var tmp = ByteArray(1)
        tmp[0] = opcode
        val brow = ByteTools.toBEByteArray(doubleVal)
        tmp = ByteTools.append(brow, tmp)
        record = tmp
    }

    override fun toString(): String {
        return string
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -1650136303920724485L
    }

}