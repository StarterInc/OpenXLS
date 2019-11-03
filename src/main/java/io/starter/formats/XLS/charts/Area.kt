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
 * **Area: Chart Group is an Area Chart Group (0x101a)**
 * 4	grbit		2	formatflags
 *
 *
 * grbit:
 * 0	0	01h	fStacked	Series in this group are stacked
 * 1	02h	f100		Each cat is broken down as a percentge
 * 2	04h	fHasShadow	1= this are has a shadow
 * 1	7-0 FFh reserved	0
 */
class Area : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    protected var fStacked = false
    protected var f100 = false
    protected var fHasShadow = false

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0)

    /**
     * @return truth of "Chart is Stacked"
     */
    override var isStacked: Boolean
        get() = fStacked
        set(bIsStacked) {
            fStacked = bIsStacked
            grbit = ByteTools.updateGrBit(grbit, fStacked, 0)
            updateRecord()
        }

    override var is100Percent: Boolean
        get() = f100
        set(bOn) {
            f100 = bOn
            grbit = ByteTools.updateGrBit(grbit, f100, 1)
            updateRecord()
        }

    /**
     * @return String XML representation of this chart-type's options
     */
    override val optionsXML: String
        get() {
            val sb = StringBuffer()
            if (this.fStacked)
                sb.append(" Stacked=\"true\"")
            if (this.f100)
                sb.append(" PercentageDisplay=\"true\"")
            if (this.fHasShadow)
                sb.append(" Shadow=\"true\"")
            return sb.toString()
        }

    override fun init() {
        super.init()
        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        fStacked = grbit and 0x1 == 0x1
        f100 = grbit and 0x2 == 0x2
        fHasShadow = grbit and 0x4 == 0x4
        chartType = ChartConstants.AREACHART    // 20070703 KSC
    }

    protected fun updateRecord() {
        grbit = ByteTools.updateGrBit(grbit, fStacked, 0)
        grbit = ByteTools.updateGrBit(grbit, f100, 1)
        grbit = ByteTools.updateGrBit(grbit, fHasShadow, 2)
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }


    /**
     * Handle setting options
     */
    override fun setChartOption(op: String, `val`: String): Boolean {
        var bHandled = false
        if (op.equals("Stacked", ignoreCase = true)) {
            this.fStacked = `val` == "true"
            bHandled = true
        } else if (op.equals("PercentageDisplay", ignoreCase = true)) {
            this.f100 = `val` == "true"
            bHandled = true
        } else if (op.equals("Shadow", ignoreCase = true)) {
            this.fHasShadow = `val` == "true"
            bHandled = true
        }
        if (bHandled)
            this.updateRecord()
        return bHandled
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -4600344312324775780L

        val prototype: XLSRecord?
            get() {
                val b = Area()
                b.opcode = XLSConstants.AREA
                b.data = b.PROTOTYPE_BYTES
                b.init()
                return b
            }
    }

}
