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
 * **RadarArea: Chart Group Is a Radar Area Chart Group(0x1040)**
 * (i.e. a filled radar chart)
 *
 *
 * A - fRdrAxLab (1 bit): A bit that specifies whether category (3) labels are displayed.
 *
 *
 * B - fHasShadow (1 bit): A bit that specifies whether the data points in the chart group have shadows.
 *
 *
 * reserved (14 bits): MUST be zero, and MUST be ignored.
 *
 *
 * unused (2 bytes): Undefined and MUST be ignored.
 */
class RadarArea : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    private var fRdrAxLab = false

    private val PROTOTYPE_BYTES = byteArrayOf(1, 0)

    /**
     * @return String XML representation of this chart-type's options
     */
    override val optionsXML: String
        get() {
            val sb = StringBuffer()
            if (fRdrAxLab)
                sb.append(" AxisLabels=\"true\"")
            return sb.toString()
        }

    override fun init() {
        super.init()
        chartType = ChartConstants.RADARAREACHART
        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        fRdrAxLab = grbit and 0x1 == 0x1
    }

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
        }
        if (bHandled)
            updateRecord()
        return bHandled
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 5731720802332312350L

        val prototype: XLSRecord?
            get() {
                val r = RadarArea()
                r.opcode = XLSConstants.RADARAREA
                r.data = r.PROTOTYPE_BYTES
                r.init()
                return r
            }
    }
}
