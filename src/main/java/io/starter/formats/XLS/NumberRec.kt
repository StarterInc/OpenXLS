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

import io.starter.OpenXLS.ExcelTools
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger


/**
 * **NUMBER: BiffRec Value, Floating-Point Number (203h)**<br></br>
 * This record stores an internal numeric type.  Stores data in one of four
 * RK 'types' which determine whether it is an integer or an IEEE floating point
 * equivalent.
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2       Row Number
 * 6       col         2       Column Number of the RK record
 * 8       ixfe        2       Index to XF cell format record
 * 10      num         8       Floating point number
</pre> *
 *
 * @see RK
 *
 * @see MULRK
 */

class NumberRec : XLSCellRecord {
    //int t;
    override var dblVal: Double = 0.toDouble()
        internal set

    /**
     * Get the value of the record as a Float.
     * Value must be parseable as an Float or it
     * will throw a NumberFormatException.
     */
    override//    	setNumberVal(d);	// original
    // 20090708 KSC: handle casting issues by converting float to string first
    var floatVal: Float
        get() = dblVal.toFloat()
        set(d) = setNumberVal(Double(d.toString()))

    override var intVal: Int
        get() {
            if (dblVal > Integer.MAX_VALUE) {
                throw NumberFormatException("Cell value is larger than the maximum java signed int size")
            }
            if (dblVal < Integer.MIN_VALUE) {
                throw NumberFormatException("Cell value is smaller than the minimum java signed int size")
            }
            return dblVal.toInt()
        }
        set(i) {
            val d = i.toDouble()
            setNumberVal(d)
        }

    override var stringVal: String?
        get() = if (isIntNumber) dblVal.toInt().toString() else ExcelTools.getNumberAsString(dblVal)
        set

    /**
     * Constructor which takes an Integer value
     */
    constructor(`val`: Int) : super() {
        opcode = XLSConstants.NUMBER
        length = 14.toShort()
        //	setLabel("NUMBER");
        setData(ByteArray(14))
        originalsize = 14
        setNumberVal(`val`.toLong())
        isIntNumber = true
        isFPNumber = false
    }

    /**
     * Constructor which takes a number value
     */
    constructor(`val`: Long) : this(`val`.toDouble()) {}

    /**
     * Constructor which takes a number value
     */
    constructor(`val`: Double) : super() {
        opcode = XLSConstants.NUMBER
        length = 14.toShort()
        //  setLabel("NUMBER");
        setData(ByteArray(14))
        originalsize = 14
        setNumberVal(`val`)
    }

    override fun init() {
        super.init()
        val l = 0
        val m = 0
        // get the row information
        super.initRowCol()
        val s = ByteTools.readShort(getByteAt(4).toInt(), getByteAt(5).toInt())
        ixfe = s.toInt()
        // get the long
        //      get the long
        dblVal = ByteTools.eightBytetoLEDouble(getBytesAt(6, 8)!!)
        isValueForCell = true
        if (DEBUGLEVEL > 5) Logger.logInfo("NumberRec: $cellAddress:$stringVal")

        val d = dblVal.toString()
        if (d.length > 12) {
            isDoubleNumber = true
            isFPNumber = true
            isIntNumber = false
        } else if (d.substring(d.length - 2) == ".0" && dblVal < Integer.MAX_VALUE) {
            // this is for io.starter.OpenXLS output files, as we put int's into number records!
            isIntNumber = true
            isFPNumber = false
            isDoubleNumber = false
        } else {
            if (dblVal < java.lang.Float.MAX_VALUE || dblVal * -1 < java.lang.Float.MAX_VALUE) {
                isFPNumber = true
                isIntNumber = false
            } else {
                // isFPNumber=true;
                isDoubleNumber = true
                isIntNumber = false
            }
        }
    }

    override fun setDoubleVal(v: Double) {
        setNumberVal(v)
    }

    constructor() : super() {}

    internal fun setNumberVal(d: Long) {
        val b: ByteArray
        val rkdata = getData()
        b = ByteTools.toBEByteArray(d.toDouble())
        System.arraycopy(b, 0, rkdata!!, 6, 8)
        setData(rkdata)
        init()
    }

    internal fun setNumberVal(d: Double) {
        val b: ByteArray
        val rkdata = getData()
        b = ByteTools.toBEByteArray(d)
        System.arraycopy(b, 0, rkdata!!, 6, 8)
        setData(rkdata)
        init()
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 7489308348300854345L
    }
}