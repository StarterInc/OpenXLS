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
package io.starter.formats.XLS.charts

import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools

/**
 * **Axcent: Axis Options(0x1062)**
 *
 *
 * 4		catMin		2		minimum date on axis.
 * If fAutoMin is set to 1, MUST be ignored.
 * If fDateAxis is set to 0, MUST be ignored.
 * 6		catMax		2		maximum date on axis.
 * fAutoMax is set to 1, MUST be ignored.
 * If fDateAxis is set to 0, MUST be ignored.
 * 8		catMajor	2		value of major unit
 * MUST be greater than or equal to catMinor when duMajor is equal to duMinor.
 * If fAutoMajor is set to 1, MUST be ignored.
 * If fDateAxis is set to 0, MUST be ignored.
 * 10		duMajor		2		Date Enumeration specifies unit of time for use of catMajor
 * If fDateAxis is set to 0, MUST be ignored.
 * 12		catMinor	2		value of minor unit
 * 14		duMinor		2		time units of minor unit
 * If fDateAxis is set to 0, MUST be ignored.
 * 16		duBase		2		smallest unit of time used by the axis.
 * If fAutoBase is set to 1, this field MUST be ignored.
 * If fDateAxis is set to 0, MUST be ignored.
 * 18		catCrossDate 2		crossing point of value axis (date)
 * 20		grbit		2
 *
 *
 * 0	0x1		fAutoMin	1= use default
 * 1	0x2		fAutoMax	""
 * 2	0x4		fAutoMajor	""
 * 3	0x8		fAutoMinor	""
 * 4	0x10	fdateAxis	1= this is a date axis
 * 5	0x20	fAutoBase	1= use default base
 * 6	0x40	fAutoCross	""
 * 7	0x80	fAutoDate	1= use default date settings for axis
 */
class Axcent : GenericChartObject(), ChartObject {
    // 20071223 KSC: add access methods
    var catMin: Short = 0
        internal set
    var catMax: Short = 0
        internal set
    var catMajor: Short = 0
        internal set
    var duMajor: Short = 0
        internal set
    var catMinor: Short = 0
        internal set
    var duMinor: Short = 0
        internal set
    var duBase: Short = 0
        internal set
    internal var catCrossDate: Short = 0
    internal var grbit: Short = 0

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, -17, 0)

    val isDefaultMin: Boolean
        get() = grbit and 0x1 == 0x1

    val isDefaultMax: Boolean
        get() = grbit and 0x2 == 0x2

    val isDefaultMajorUnits: Boolean
        get() = grbit and 0x4 == 0x4

    val isDefaultMinorUnits: Boolean
        get() = grbit and 0x8 == 0x8

    override fun init() {
        super.init()
        // 20071223 KSC: Start parsing of options
        catMin = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        catMax = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        catMajor = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        duMajor = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        catMinor = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())
        duMinor = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt())
        duBase = ByteTools.readShort(this.getByteAt(12).toInt(), this.getByteAt(13).toInt())
        catCrossDate = ByteTools.readShort(this.getByteAt(14).toInt(), this.getByteAt(15).toInt())
        grbit = ByteTools.readShort(this.getByteAt(16).toInt(), this.getByteAt(17).toInt())
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -660100252646337769L

        // 20070723 KSC: Need to create new records
        // important when we parse these options ...
        val prototype: XLSRecord?
            get() {
                val a = Axcent()
                a.opcode = XLSConstants.AXCENT
                a.data = a.PROTOTYPE_BYTES
                a.init()
                return a
            }
    }
}
