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
 * **Serfmt: Series Format(0x105d)**
 *
 *
 * Specifies series formatting information
 *
 *
 * 0	grbit		2
 *
 *
 * bits
 * 0		0x1	fSmoothedLine		1= the line series has a smoothed line (Line, Scatter or Radar)
 * 1		0x2	f3DBubbles			1= draw bubbles with 3-D effects
 * 2		0x4 fArShadow			1= specifies whether the data markers are displayed with a shadow on bubble,
 * scatter, radar, stock, and line chart groups.
 * rest are reserved
 */
class Serfmt : GenericChartObject(), ChartObject {
    // 20070810 KSC: parse options ...
    private var grbit: Short = 0
    var smoothLine = false
        private set
    var 3DBubbles = false
    private set
    /**
     * data markers are displayed with a shadow on bubble,
     * scatter, radar, stock, and line chart groups.
     */
    var shadow = false
        private set

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0)

    /**
     * @return String XML representation of this chart-type's options
     */
    override val optionsXML: String
        get() {
            val sb = StringBuffer()
            if (smoothLine)
                sb.append(" SmoothedLine=\"true\"")
            if (3 DBubbles)
                sb.append(" ThreeDBubbles=\"true\"")
            if (shadow)
                sb.append(" ArShadow=\"true\"")
            return sb.toString()
        }

    override fun init() {
        super.init()
        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        smoothLine = grbit and 0x1 == 0x1
        3 DBubbles =(grbit and 0x2) == 0x2
        shadow = grbit and 0x4 == 0x4
    }

    private fun updateRecord() {
        grbit = ByteTools.updateGrBit(grbit, smoothLine, 0)
        grbit = ByteTools.updateGrBit(grbit, 3 DBubbles, 1)
        grbit = ByteTools.updateGrBit(grbit, shadow, 2)
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    /**
     * Handle setting options from XML in a generic manner
     */
    override fun setChartOption(op: String, `val`: String): Boolean {
        var bHandled = false
        if (op.equals("SmoothedLine", ignoreCase = true)) {
            smoothLine = `val` == "true"
            bHandled = true
        } else if (op.equals("ThreeDBubbles", ignoreCase = true)) {
            3 DBubbles = `val` == "true"
            bHandled = true
        } else if (op.equals("ArShadow", ignoreCase = true)) {
            shadow = `val` == "true"
            bHandled = true
        }
        if (bHandled)
            updateRecord()
        return bHandled
    }


    fun setHas3dBubbles(has3dBubbles: Boolean) {
        3 DBubbles = has3dBubbles
                updateRecord()
    }

    /**
     * sets whether the parent chart or series has smoothed lines
     *
     * @param b
     */
    fun setSmoothedLine(b: Boolean) {
        smoothLine = b
        updateRecord()
    }

    /**
     * data markers are displayed with a shadow on bubble,
     * scatter, radar, stock, and line chart groups.
     *
     * @param b
     */
    fun setHasShadow(b: Boolean) {
        shadow = b
        updateRecord()
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -8307035373276421283L

        val prototype: XLSRecord?
            get() {
                val s = Serfmt()
                s.opcode = XLSConstants.SERFMT
                s.data = s.PROTOTYPE_BYTES
                s.init()
                return s
            }
    }

}
