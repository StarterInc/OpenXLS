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
 * **Pie: Chart Group Is a pie Chart Group(0x1019)**
 *
 *
 * 4	anStart		2		Angle of the first pie slice expressed in degrees.  Must be <= 360
 * 6	pcDonut		2		0= true pie chart, non-zero= size of center hole in a donut chart, as a percentage
 * 8	grbit		2		Option Flags
 *
 *
 *
 *
 * grbit:
 * 0	0x1		fHasShadow		1= has shadow
 * 1	0x2		fShowLdrLines	1= show leader lines to data labels
 */
class Pie : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    private var fHasShadow = false
    private var fShowLdrLines = false
    var donutPercentage: Short = 0
        protected set
    protected var anStart: Short = 0

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 0, 0, 2, 0)

    /**
     * return the Doughnut hole size (if > 0)
     *
     * @return
     */
    /**
     * size of center hole in a donut chart, as a percentage.  0 for pie
     *
     * @return
     */
    var doughnutSize: Int
        get() = donutPercentage.toInt()
        set(s) {
            donutPercentage = s.toShort()
            val b = ByteTools.shortToLEBytes(donutPercentage)
            this.data[2] = b[0]
            this.data[3] = b[1]
        }

    /**
     * @return String XML representation of this chart-type's options
     */
    override val optionsXML: String
        get() {
            val sb = StringBuffer()
            if (fHasShadow)
                sb.append(" Shadow=\"true\"")
            if (fShowLdrLines)
                sb.append(" ShowLdrLines=\"true\"")
            if (donutPercentage > 0)
                sb.append(" Donut=\"$donutPercentage\"")
            return sb.toString()
        }

    override fun init() {
        super.init()
        val data = data
        anStart = ByteTools.readShort(data!![0].toInt(), data[1].toInt())
        donutPercentage = data[2].toShort()
        if (donutPercentage.toInt() == 0)
            chartType = ChartConstants.PIECHART // 20070703 KSC
        else
            chartType = ChartConstants.DOUGHNUTCHART
        grbit = ByteTools.readShort(data[4].toInt(), data[5].toInt())
        fHasShadow = grbit and 0x1 == 0x1
        fShowLdrLines = grbit and 0x2 == 0x2
    }

    // 20070703 KSC and below ...
    private fun updateRecord() {
        this.data[2] = donutPercentage.toByte()
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[4] = b[0]
        this.data[5] = b[1]
    }

    fun setAsPieChart() {
        donutPercentage = 0
        chartType = ChartConstants.PIECHART
        this.updateRecord()
    }

    fun setAsDoughnutChart() {
        donutPercentage = 50    // default %
        chartType = ChartConstants.DOUGHNUTCHART
        this.updateRecord()
    }

    /**
     * return the Angle of the first pie slice expressed in degrees.  Must be <= 360
     *
     * @return
     */
    fun getAnStart(): Int {
        return anStart.toInt()
    }


    /**
     * sets the Angle of the first pie slice expressed in degrees.  Must be <= 360
     *
     * @return
     */
    fun setAnStart(a: Int) {
        anStart = a.toShort()
        val b = ByteTools.shortToLEBytes(anStart)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    /**
     * Handle setting options from XML in a generic manner
     */
    override fun setChartOption(op: String, `val`: String): Boolean {
        var bHandled = false
        if (op.equals("Shadow", ignoreCase = true)) {
            setHasShadow(true)
            bHandled = true
        } else if (op.equals("ShowLdrLines", ignoreCase = true)) {
            setShowLdrLines(true)
            bHandled = true
        } else if (op.equals("Donut", ignoreCase = true)) {
            setDonutPercentage(`val`)
            bHandled = true
        } else if (op.equals("donutSize", ignoreCase = true)) {
            setDonutPercentage(`val`)
            bHandled = true
        }
        return bHandled
    }

    override fun getChartOption(op: String): String? {
        return if (op == "ShowLdrLines") { // Pie
            this.showLdrLines().toString()
        } else if (op == "donutSize") { // Pie
            this.donutPercentage.toString()
        } else
            super.getChartOption(op)
    }

    override fun hasShadow(): Boolean {
        return fHasShadow
    }

    fun showLdrLines(): Boolean {
        return fShowLdrLines
    }

    fun setHasShadow(bHasShadow: Boolean) {
        fHasShadow = bHasShadow
        grbit = ByteTools.updateGrBit(grbit, fHasShadow, 0)
        updateRecord()
    }

    fun setShowLdrLines(bShowLdrLines: Boolean) {
        fShowLdrLines = bShowLdrLines
        grbit = ByteTools.updateGrBit(grbit, fShowLdrLines, 1)
        updateRecord()
    }

    fun setDonutPercentage(`val`: String) {
        try {
            donutPercentage = java.lang.Short.valueOf(`val`).toShort()
        } catch (e: Exception) {
        }

        updateRecord()
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -7851320124576950635L

        val prototype: XLSRecord?
            get() {
                val b = Pie()
                b.opcode = XLSConstants.PIE
                b.data = b.PROTOTYPE_BYTES
                b.init()
                return b
            }
    }

}
