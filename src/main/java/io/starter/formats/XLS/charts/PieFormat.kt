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
 * **PieFormat: Position of the Pie Slice(0x100b)**
 *
 *
 * percentage		2		distance of pie slice from center of pie as %
 */
class PieFormat : GenericChartObject(), ChartObject {
    internal var percentage: Short = 0

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0)

    override val optionsXML: String
        get() = " Percentage=\"$percentage\""

    override fun init() {
        super.init()
        this.data
        percentage = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())

    }

    // 20070717 KSC: get/set methods for format options
    fun getPercentage(): Short {
        return percentage
    }

    fun setPercentage(p: Short) {
        percentage = p
        updateRecord()
    }

    private fun updateRecord() {
        val b = ByteTools.shortToLEBytes(percentage)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    override fun toString(): String {
        return "PieFormat:  Percentage=$percentage"
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 28305957039802849L

        // 20070716 KSC: Need to create new records
        val prototype: XLSRecord?
            get() {
                val pf = PieFormat()
                pf.opcode = XLSConstants.PIEFORMAT
                pf.data = pf.PROTOTYPE_BYTES
                pf.init()
                return pf
            }
    }
}
