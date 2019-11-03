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
 * **AxisLineFormat:  Defines a Line that spans an Axis (0x1021)**
 * The AxisLineFormat  record specifies which part of the axis is specified by the
 * LineFormat record that follows.
 * <br></br>
 * 4		id		2		Axis Line identifier:
 * 0= the axis line itself,
 * 1= major grid line,
 * 2= minor grid line,
 * 3= walls or floor
 * if 3, MUST be preceded by an Axis record with the wType set to:
 * 0x0000	The walls of a 3-D chart
 * 0x0001	The floor of a 3-D chart
 */
class AxisLineFormat : GenericChartObject(), ChartObject {
    var id: Short = 0
        private set

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0)

    override fun init() {
        super.init()
        id = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
    }

    fun setId(id: Int) {
        this.id = id.toShort()
        val b = ByteTools.shortToLEBytes(this.id)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    override fun toString(): String {
        return "AxisLineFormat: " + if (id.toInt() == 0) "Axis" else if (id.toInt() == 1) "Major" else if (id.toInt() == 2) "Minor" else if (id.toInt() == 3) "Wall or Floor" else "Unknown"
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 5243346695500373630L

        //  id MUST be greater than the id field values in preceding AxisLine records in the current axis
        val prototype: XLSRecord?
            get() {
                val a = AxisLineFormat()
                a.opcode = XLSConstants.AXISLINEFORMAT
                a.data = a.PROTOTYPE_BYTES
                a.init()
                return a
            }
        val ID_AXIS_LINE = 0
        val ID_MAJOR_GRID = 1
        val ID_MINOR_GRID = 2
        val ID_WALLORFLOOR = 3
    }

}
