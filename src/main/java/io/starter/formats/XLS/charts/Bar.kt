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
 * Bar: Chart Group is a Bar or Column Chart Group (0x1017)
 * NOTE: Bar is also base type for Pyramid, Cone and Cylinder charts;
 * actual chart type is determined also by bar shape
 * see ChartFormat.getChartType for more information
 *
 *
 *
 *
 * 4	pcOverlap	2	Space between bars (default= 0)
 * values: 	-100 to -1		Size of the separation between data points
 * 0				No overlap.
 * 1 to 100		Size of the overlap between data points
 * 6	pcGap		2	Space between categories (%) (default=50%)
 * An unsigned integer that specifies the width of the gap between the categories and the left and right edges of the plot area
 * as a percentage of the data point width divided by 2. It also specifies the width of the gap between adjacent categories
 * as a percentage of the data point width. MUST be less than or equal to 500.
 * 8	grbit		2
 *
 *
 * grbit:
 *
 *
 * 0	0	0x1		fTranspose		1= horizontal bars (=bar chart)	0= vertical bars (= column chart)
 * 1	0x2		fStacked		Stack the displayed values
 * 2	0x4		f100			Each category is displayed as a percentage
 * 3	0x8		fHasShadow		1= this bar has a shadow
 */
class Bar : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    var fStacked = false
    var f100 = false
    var fHasShadow = false
    var pcOverlap: Short = 0
    var pcGap: Short = 50

    // 20070716 KSC: get/set methods for format options
    /**
     * sets this bar/col chart to have stacked series:  series are drawn next to each other along the value axis
     *
     * @param bIsClustered
     */
    override var isStacked: Boolean
        get() = fStacked
        set(bIsStacked) {
            fStacked = bIsStacked
            grbit = ByteTools.updateGrBit(grbit, fStacked, 1)
            if (bIsStacked) {
                pcOverlap = -100
                pcGap = 150
            }
            updateRecord()
        }

    override var is100Percent: Boolean
        get() = f100
        set(bOn) {
            f100 = bOn
            grbit = ByteTools.updateGrBit(grbit, f100, 2)
            if (bOn) {
                pcOverlap = -100
                pcGap = 150
            }
            updateRecord()
        }

    val gap: Int
        get() = pcGap.toInt()

    val overlap: Int
        get() = pcOverlap.toInt()

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, -106, 0, 0, 0)

    override fun init() {
        super.init()
        pcOverlap = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        // should be 50 default, but seems to be 150 ??????
        pcGap = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        // pcGap: 0x0096 specifies that the width of the gap between adjacent categories is 150% of the data point width. It also specifies that the width of the gap between the categories (3) and the left and right edges of the plot area is 75% of the data point width.
        grbit = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        if (grbit and 0x1 == 0x1) {
            chartType = ChartConstants.BARCHART
        } else {
            chartType = ChartConstants.COLCHART
        }
        fStacked = grbit and 0x2 == 0x2
        f100 = grbit and 0x4 == 0x4
        fHasShadow = grbit and 0x8 == 0x8
    }

    override fun hasShadow(): Boolean {
        return fHasShadow
    }

    fun setHasShadow(bHasShadow: Boolean) {
        fHasShadow = bHasShadow
        grbit = ByteTools.updateGrBit(grbit, fHasShadow, 3)
        updateRecord()
    }

    /**
     * sets this bar/col chart to have clustered series:  series are drawn next to each other along the category axis
     *
     * @param bIsClustered
     */
    fun setIsClustered() {
        isStacked = false
        is100Percent = false
    }

    /**
     * sets the Space between bars (default= 0)
     *
     * @param overlap
     */
    fun setpcOverlap(overlap: Int) {
        pcOverlap = overlap.toShort()
        updateRecord()
    }

    /**
     * sets the Space between categories (%) (default=50%)
     *
     * @param gap
     */
    fun setpcGap(gap: Int) {
        pcGap = gap.toShort()
        updateRecord()
    }

    fun setAsBarChart() {
        grbit = ByteTools.updateGrBit(grbit, true, 0)    // set 0'th bit
        chartType = ChartConstants.BARCHART
        this.updateRecord()
    }

    fun setAsColumnChart() {
        grbit = ByteTools.updateGrBit(grbit, false, 0)    // clear 0'th bit
        chartType = ChartConstants.COLCHART
        this.updateRecord()
    }

    private fun updateRecord() {
        var b = ByteTools.shortToLEBytes(pcOverlap)
        this.data[0] = b[0]
        this.data[1] = b[1]
        b = ByteTools.shortToLEBytes(pcGap)
        this.data[2] = b[0]
        this.data[3] = b[1]
        b = ByteTools.shortToLEBytes(grbit)
        this.data[4] = b[0]
        this.data[5] = b[1]
    }

    /**
     * Set specific options
     */
    override fun setChartOption(op: String, `val`: String): Boolean {
        var bHandled = false
        if (op.equals("Stacked", ignoreCase = true)) {
            isStacked = true
            bHandled = true
        } else if (op.equals("PercentageDisplay", ignoreCase = true)) {
            is100Percent = true
            bHandled = true
        } else if (op.equals("Shadow", ignoreCase = true)) {
            setHasShadow(true)
            bHandled = true
        } else if (op.equals("Overlap", ignoreCase = true)) {
            setpcOverlap(Integer.parseInt(`val`))
            bHandled = true
        } else if (op.equals("Gap", ignoreCase = true)) {
            setpcGap(Integer.parseInt(`val`))
            bHandled = true
        }
        return bHandled
    }

    /**
     * look up bar-specific chart options such as "Gap" or "Overlap" setting
     */
    override fun getChartOption(op: String): String? {
        return if (op == "Gap") { // Bar
            this.gap.toString()
        } else if (op == "Overlap") { // Bar
            //    		return String.valueOf(Math.abs(this.getOverlap()));		// KSC: TESTING:  OOXML apparently needs +100 pcOverlap NOT -100 ... WHY and TRUE FOR ALL CASES?????
            this.overlap.toString()        // KSC: TESTING:  OOXML apparently needs +100 pcOverlap NOT -100 ... WHY and TRUE FOR ALL CASES?????
        } else
            super.getChartOption(op)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 8917510368688674273L

        val prototype: XLSRecord?
            get() {
                val b = Bar()
                b.opcode = XLSConstants.BAR
                b.data = b.PROTOTYPE_BYTES
                b.init()
                b.setAsColumnChart()
                return b
            }
    }
}
