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
 * **Line: Chart Group Is a Line Chart Group (0x1018)**
 *
 *
 * 4	grbit		2		flags
 *
 *
 * 0		0x1		fStacked		Stack the displayed values
 * 1		0x2		f100			Each category is broken down as a percentage
 * 2		0x4		fHasShadow		1= this line has a shadow
 */
class Line : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    private var fStacked = false
    private var f100 = false
    private var fHasShadow = false

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0)

    /**
     * @return String XML representation of this chart-type's options
     */
    override val optionsXML: String
        get() {
            val sb = StringBuffer()
            if (fStacked)
                sb.append(" Stacked=\"true\"")
            if (f100)
                sb.append(" PercentageDisplay=\"true\"")
            if (fHasShadow)
                sb.append(" Shadow=\"true\"")
            return sb.toString()
        }

    override var is100Percent: Boolean
        get() = super.is100Percent
        set(bOn) {
            f100 = bOn
            grbit = ByteTools.updateGrBit(grbit, f100, 1)
            updateRecord()
        }

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

    override fun init() {
        super.init()
        chartType = ChartConstants.LINECHART    // 20070703 KSC
        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        fStacked = grbit and 0x1 == 0x1
        f100 = grbit and 0x2 == 0x2
        fHasShadow = grbit and 0x4 == 0x4

    }

    // 20070725 KSC:
    private fun updateRecord() {
        grbit = ByteTools.updateGrBit(grbit, fStacked, 0)
        grbit = ByteTools.updateGrBit(grbit, f100, 1)
        grbit = ByteTools.updateGrBit(grbit, fHasShadow, 2)
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    fun setAsStockChart() {
        chartType = ChartConstants.STOCKCHART
    }

    /**
     * Handle setting options from XML in a generic manner
     */
    override fun setChartOption(op: String, `val`: String): Boolean {
        var bHandled = false
        if (op.equals("Stacked", ignoreCase = true)) {
            fStacked = `val` == "true"
            bHandled = true
        } else if (op.equals("PercentageDisplay", ignoreCase = true)) {
            f100 = `val` == "true"
            bHandled = true
        } else if (op.equals("Shadow", ignoreCase = true)) {
            fHasShadow = `val` == "true"
            bHandled = true
        } else if (op.equals("Smooth", ignoreCase = true)) {    // smooth lines
            /*            if (b instanceof Serfmt) {
            	Serfmt s= (Serfmt) b;
            	if (s.getSmoothLine()) {*/
        }
        if (bHandled)
            updateRecord()
        return bHandled
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 6526476252082906554L

        val prototype: XLSRecord?
            get() {
                val b = Line()
                b.opcode = XLSConstants.LINE
                b.data = b.PROTOTYPE_BYTES
                b.init()
                return b
            }
    }
}
