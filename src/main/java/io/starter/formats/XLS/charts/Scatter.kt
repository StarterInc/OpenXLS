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
 * **Scatter: Chart Group is a Scatter Chart Group(0x101b)**
 *
 *
 * 4	pcBubbleSizeRatio	2		Percent of largest bubble compared to chart in general default= 100
 * 6	wBubbleSize			2		Bubble size: 1= bubble size is area, 2= bubble size is width	default= 1
 * 8	grbit				2		flags
 *
 *
 * grbit
 * 0	0x1		fBubbles		1= this is a bubble series
 * 1	0x2		fShowNegBubbles	1= show negative bubbles
 * 2	0x4		fHasShadow		1= bubble series has a shadow
 */
class Scatter : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    private var pcBubbleSizeRatio: Short = 100
    private var wBubbleSize: Short = 1
    private var fBubbles = false
    private var fShowNegBubbles = false
    private var fHasShadow = false

    private val PROTOTYPE_BYTES = byteArrayOf(100, 0, 1, 0, 0, 0)

    /**
     * @return String XML representation of this chart-type's options
     */
    override val optionsXML: String
        get() {
            val sb = StringBuffer()
            if (pcBubbleSizeRatio.toInt() != 100)
                sb.append(" BubbleSizeRatio=\"$pcBubbleSizeRatio\"")
            if (wBubbleSize.toInt() != 1)
                sb.append(" BubbleSize=\"$wBubbleSize\"")
            if (fShowNegBubbles)
                sb.append(" ShowNeg=\"true\"")
            if (fHasShadow)
                sb.append(" Shadow=\"true\"")
            return sb.toString()
        }

    override fun init() {
        super.init()
        // 20070703 KSC:
        pcBubbleSizeRatio = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        wBubbleSize = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        grbit = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        fBubbles = grbit and 0x1 == 0x1
        fShowNegBubbles = grbit and 0x2 == 0x2
        fHasShadow = grbit and 0x4 == 0x4
        if (fBubbles)
            chartType = ChartConstants.BUBBLECHART
        else
            chartType = ChartConstants.SCATTERCHART
    }

    // 20070703 KSC
    private fun updateRecord() {
        grbit = ByteTools.updateGrBit(grbit, fBubbles, 0)
        grbit = ByteTools.updateGrBit(grbit, fShowNegBubbles, 1)
        grbit = ByteTools.updateGrBit(grbit, fHasShadow, 2)
        var b = ByteTools.shortToLEBytes(pcBubbleSizeRatio)
        this.data[0] = b[0]
        this.data[1] = b[1]
        b = ByteTools.shortToLEBytes(wBubbleSize)
        this.data[2] = b[0]
        this.data[3] = b[1]
        b = ByteTools.shortToLEBytes(grbit)
        this.data[4] = b[0]
        this.data[5] = b[1]
    }

    fun setAsScatterChart() {
        fBubbles = false
        chartType = ChartConstants.SCATTERCHART
        this.updateRecord()
    }

    fun setAsBubbleChart() {
        fBubbles = true
        chartType = ChartConstants.BUBBLECHART
        this.updateRecord()
    }

    /**
     * Handle setting options from XML in a generic manner
     */
    override fun setChartOption(op: String, `val`: String): Boolean {
        var bHandled = false
        if (op.equals("BubbleSizeRatio", ignoreCase = true)) {
            pcBubbleSizeRatio = java.lang.Short.parseShort(`val`)
            bHandled = true
        } else if (op.equals("BubbleSize", ignoreCase = true)) {
            wBubbleSize = java.lang.Short.parseShort(`val`)
            bHandled = true
        } else if (op.equals("ShowNeg", ignoreCase = true)) {
            fShowNegBubbles = `val` == "true"
            bHandled = true
        } else if (op.equals("Shadow", ignoreCase = true)) {
            fHasShadow = `val` == "true"
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
        private val serialVersionUID = -289334135036242100L

        val prototype: XLSRecord?
            get() {
                val b = Scatter()
                b.opcode = XLSConstants.SCATTER
                b.data = b.PROTOTYPE_BYTES
                b.init()
                b.setAsScatterChart()
                return b
            }
    }

}	
