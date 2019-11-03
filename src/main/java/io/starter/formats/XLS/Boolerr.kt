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
package io.starter.formats.XLS

import io.starter.toolkit.ByteTools


/**
 * **Boolerr: BiffRec Value, Boolean or Error (0x205)**<br></br>
 * Describes a cell that contains a constant Boolean or error
 * value.
 *
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2       Row
 * 6       col         2       Column
 * 8       ixfe        2       Index to XF record
 * 10      bBoolErr    1       Boolean value or error value
 * 11      fError      1       Boolean/error flag (0 = Boolean, 1 = Error)
 *
 * Boolean vals are 1 = true, 0 = false.
 *
</pre> *
 */

class Boolerr : XLSCellRecord() {
    private var `val`: Boolean = false
    /**
     * Returns whether the value is a Boolean or
     * an error.
     */
    internal var isErr = false
        private set
    private var errorval: Byte = 0

    /**
     * get the int val
     */
    override var intVal: Int
        get() = if (this.booleanVal) 1 else 0
        set

    /**
     * return boolean value in float version 0 or 1
     */
    override val floatVal: Float
        get() = if (this.booleanVal) 1f else 0f

    /**
     * return the boolean value in double version 0 or 1
     */
    override val dblVal: Double
        get() = if (this.booleanVal) 1.0 else 0.0

    /**
     * Returns the valid error code for this valrec
     */
    val errorCode: String
        get() {
            if (!this.isErr) {
                return isErr.toString()
            }
            var retval = ""
            if (errorval and 0x0 == 0x0) retval = "#NULL!"
            if (errorval and 0x7 == 0x7) retval = "#DIV/0!"
            if (errorval and 0xF == 0xF) retval = "#VALUE!"
            if (errorval and 0x17 == 0x17) retval = "#REF!"
            if (errorval and 0x1D == 0x1D) retval = "#NAME?"
            if (errorval and 0x24 == 0x24) retval = "#NUM!"
            if (errorval and 0x2A == 0x2A) retval = "#N/A"
            return retval
        }

    /**
     * get the String val
     */
    override// char encoding of true/false should be irrelevant -
    //! NOPE, 'cause this also has error codes in it -NR 10/03
    var stringVal: String?
        get() = this.getStringVal(null)
        set

    /**
     * Get the value of the record as a Boolean.
     * Value must be parseable as a Boolean.
     */
    override var booleanVal: Boolean
        get() = `val`
        set(newv) {
            if (newv)
                this.getData()[6] = 1
            else
                this.getData()[6] = 0
            this.`val` = newv
        }

    // these bytes are from a simple chart, 2 ranges a1:a2, b1:b2 - all default.  Likely will need to be modified
    // when we figure out wtf.
    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 0, 0, 21, 0, 0, 0)

    /**
     * get the String val
     */
    override fun getStringVal(encoding: String?): String? {
        return if (this.isErr) {
            this.errorCode
        } else `val`.toString()
    }

    override fun init() {
        super.init()
        // get the row information
        rw = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        colNumber = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        ixfe = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt()).toInt()
        // get the value
        var num = this.getByteAt(6).toInt()
        if (num == 0) `val` = false
        if (num == 1) `val` = true
        num = this.getByteAt(7).toInt()
        if (num == 0) isErr = false
        if (num == 1) isErr = true
        errorval = this.getByteAt(6)
        this.isValueForCell = true
        if (!isErr) {
            isBoolean = true
        } else {
            isString = true
        }
    }

    companion object {


        private val serialVersionUID = 39663492256953223L

        val prototype: XLSRecord?
            get() {
                val be = Boolerr()
                be.opcode = XLSConstants.BOOLERR
                be.setData(be.PROTOTYPE_BYTES)
                be.init()
                return be
            }
    }
}