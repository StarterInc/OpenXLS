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
 * **Radar: Chart Group Is a Radar Chart Group(0x103e)**
 *
 *
 * 4		grbit		2
 *
 *
 * 0		0x1		fRdrAxLab			1= chart contains radar axis labels
 * 1		0x2		fHasShadow			1= this radar series has a shadow
 */
class Radar : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    private var fRdrAxLab = true
    private var fHasShadow = false


    private val PROTOTYPE_BYTES = byteArrayOf(1, 0, 18, 0)

    /**
     * @return String XML representation of this chart-type's options
     */
    override val optionsXML: String
        get() {
            val sb = StringBuffer()
            if (fRdrAxLab)
                sb.append(" AxisLabels=\"true\"")
            if (fHasShadow)
                sb.append(" Shadow=\"true\"")
            return sb.toString()
        }

    override fun init() {
        super.init()
        chartType = ChartConstants.RADARCHART
        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        fRdrAxLab = grbit and 0x1 == 0x1
        fHasShadow = grbit and 0x2 == 0x2
    }

    // 20070703 KSC: taken from Bar.java
    private fun updateRecord() {
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    /**
     * Handle setting options from XML in a generic manner
     */
    override fun setChartOption(op: String, `val`: String): Boolean {
        var bHandled = false
        if (op.equals("AxisLabels", ignoreCase = true)) {
            fRdrAxLab = `val` == "true"
            grbit = ByteTools.updateGrBit(grbit, fRdrAxLab, 0)
            bHandled = true
        } else if (op.equals("Shadow", ignoreCase = true)) {
            fHasShadow = `val` == "true"
            grbit = ByteTools.updateGrBit(grbit, fHasShadow, 1)
            bHandled = true
        }
        if (bHandled)
            updateRecord()
        return bHandled
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 3443368503725347910L

        val prototype: XLSRecord?
            get() {
                val b = Radar()
                b.opcode = XLSConstants.RADAR
                b.data = b.PROTOTYPE_BYTES
                b.init()
                return b
            }
    }


}
