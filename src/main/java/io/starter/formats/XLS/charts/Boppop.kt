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
 * **Boppop: Bar of Pie/Pie of Pie chart options(0x1061)**
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 * pst (1 byte): An unsigned integer that specifies whether this chart group is a bar of pie chart group or a pie of pie chart group. MUST be a value from the following table:
 * 1= pie of pie 2= bar of pie
 * fAutoSplit (1 byte): A Boolean that specifies whether the split point of the chart group is determined automatically.
 * If the value is 1, when a bar of pie chart group or pie of pie chart group is initially created the data points from the primary pie are selected
 * and inserted into the secondary bar/pie automatically.
 * split (2 bytes): An unsigned integer that specifies what determines the split between the primary pie and the secondary bar/pie.
 * MUST be ignored if fAutoSplit is set to 1. MUST be a value from the following table:
 *
 *
 * Value	 	Type of split		   	 Meaning
 * 0x0000	    Position			    The data is split based on the position of the data point in the series as specified by iSplitPos.
 * 0x0001	    Value				    The data is split based on a threshold value as specified by numSplitValue.
 * 0x0002	    Percent				    The data is split based on a percentage threshold and the data point values represented as a percentage as specified by pcSplitPercent.
 * 0x0003	    Custom				    The data is split as arranged by the user. Custom split is specified in a following BopPopCustom record.
 *
 *
 * iSplitPos (2 bytes): A signed integer that specifies how many data points are contained in the secondary bar/pie.
 * Data points are contained in the secondary bar/pie starting from the end of the series.
 * For example, if the value is 2, the last 2 data points in the series are contained in the secondary bar/pie.
 * MUST be a value greater than or equal to 0 and less than or equal to 32000.
 * If the value is more than the number of data points in the series, the entire series will be in the secondary bar/pie, except for the first data point.
 * If split is not set to 0x0000 or fAutoSplit is set to 1, this value MUST be ignored.
 *
 *
 * pcSplitPercent (2 bytes): A signed integer that specifies the percentage below which each data point is contained in the secondary bar/pie as opposed to the primary pie.
 * The percentage value of a data point is calculated using the following formula:
 * (value of the data point x 100) / sum of all data points in the series
 * If split is not set to 0x0002 or if fAutoSplit is set to 1, this value MUST be ignored
 *
 *
 * pcPie2Size (2 bytes): A signed integer that specifies the size of the secondary bar/pie as a percentage of the size of the primary pie.
 * MUST be a value greater than or equal to 5 and less than or equal to 200.
 *
 *
 * pcGap (2 bytes): A signed integer that specifies the distance between the primary pie and the secondary bar/pie.
 * The distance is specified as a percentage of the average width of the primary pie and secondary bar/pie.
 * MUST be a value greater than or equal to 0 and less than or equal to 500,
 * where 0 is 0% of the average width of the primary pie and the secondary bar/pie, and 500 is 250% of the average width of the primary pie and the secondary bar/pie.
 *
 *
 * numSplitValue (8 bytes): An Xnum value that specifies the split when the split field is set to 0x0001
 * . The value of this field specifies the threshold that selects which data points of the primary pie move to the secondary bar/pie.
 * The secondary bar/pie contains any data points with a value less than the value of this field. If split is not set to 0x0001 or if fAutoSplit is set to 1, this value MUST be ignored.
 *
 *
 * A - fHasShadow (1 bit): A bit that specifies whether one or more data points in the chart group have shadows.
 *
 *
 * reserved (15 bits): MUST be zero, and MUST be ignored.
 */
class Boppop : GenericChartObject(), ChartObject {
    internal var fAutoSplit: Boolean = false
    internal var fHasShadow: Boolean = false
    internal var split: Short = 0
    internal var iSplitPos: Short = 0
    internal var pcSplitPercent: Short = 0
    internal var pcPieSize: Short = 0
    internal var pcGap: Short = 0
    internal var numSplitValue: Float = 0.toFloat()
    internal var pst: Byte = 0

    private val PROTOTYPE_BYTES = byteArrayOf(1, 1, 0, 0, 0, 0, 0, 0, 75, 0, 100, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    /**
     * returns true of pie of pie, false if bar of pie
     *
     * @return
     */
    var isPieOfPie: Boolean
        get() = pst.toInt() == 1
        set(b) {
            if (b)
                pst = 1
            else
                pst = 2
            this.data[0] = pst
        }

    /**
     * returns the size of the secondary bar/pie as a percentage of the size of the primary pie.
     */
    /**
     * specifies the size of the secondary bar/pie as a percentage of the size of the primary pie.
     *
     * @param s
     */
    var secondPieSize: Int
        get() = pcPieSize.toInt()
        set(s) {
            pcPieSize = s.toShort()
            val b = ByteTools.shortToLEBytes(pcPieSize)
            this.data[8] = b[0]
            this.data[9] = b[1]
        }

    /**
     * returns split type, or -1 if autosplit
     * 0x0000	    Position			    The data is split based on the position of the data point in the series as specified by iSplitPos.
     * 0x0001	    Value				    The data is split based on a threshold value as specified by numSplitValue.
     * 0x0002	    Percent				    The data is split based on a percentage threshold and the data point values represented as a percentage as specified by pcSplitPercent.
     * 0x0003	    Custom				    The data is split as arranged by the user. Custom split is specified in a following BopPopCustom record.
     *
     * @return
     */
    /**
     * specifies what determines the split between the primary pie and the secondary bar/pie.
     * 0x0000	    Position			    The data is split based on the position of the data point in the series as specified by iSplitPos.
     * 0x0001	    Value				    The data is split based on a threshold value as specified by numSplitValue.
     * 0x0002	    Percent				    The data is split based on a percentage threshold and the data point values represented as a percentage as specified by pcSplitPercent.
     * 0x0003	    Custom				    The data is split as arranged by the user. Custom split is specified in a following BopPopCustom record.
     *
     * @param t
     */
    var splitType: Int
        get() = if (fAutoSplit) -1 else split.toInt()
        set(t) {
            fAutoSplit = false
            this.data[1] = 0
            split = t.toShort()
            val b = ByteTools.shortToLEBytes(split)
            this.data[2] = b[0]
            this.data[3] = b[1]
        }

    // default
    val splitTypeOOXML: String
        get() {
            if (fAutoSplit)
                return "auto"
            when (split) {
                0 -> return "pos"
                1 -> return "val"
                2 -> return "percent"
                3 -> return "cust"
            }
            return "auto"
        }

    /**
     * returns how many data points are contained in the secondary bar/pie.
     * Data points are contained in the secondary bar/pie starting from the end of the series.
     * For example, if the value is 2, the last 2 data points in the series are contained in the secondary bar/pie.
     */
    /**
     * specifies how many data points are contained in the secondary bar/pie.
     * Data points are contained in the secondary bar/pie starting from the end of the series.
     * For example, if the value is 2, the last 2 data points in the series are contained in the secondary bar/pie.
     *
     * @param sp
     */
    var splitPos: Int
        get() = iSplitPos.toInt()
        set(sp) {
            iSplitPos = sp.toShort()
            val b = ByteTools.shortToLEBytes(iSplitPos)
            this.data[4] = b[0]
            this.data[5] = b[1]
        }

    override fun init() {
        super.init()
        val data = this.data
        pst = data!![0]
        fAutoSplit = data[1].toInt() == 1
        split = ByteTools.readShort(data[2].toInt(), data[3].toInt())
        iSplitPos = ByteTools.readShort(data[4].toInt(), data[5].toInt())
        pcSplitPercent = ByteTools.readShort(data[6].toInt(), data[7].toInt())
        pcPieSize = ByteTools.readShort(data[8].toInt(), data[9].toInt())
        pcGap = ByteTools.readShort(data[10].toInt(), data[11].toInt())
        numSplitValue = ByteTools.eightBytetoLEDouble(this.getBytesAt(12, 8)!!).toFloat()
        fHasShadow = data[20] and 1 == 1
        chartType = ChartConstants.OFPIECHART
    }

    /**
     * specifies the distance between the primary pie and the secondary bar/pie.
     * The distance is specified as a percentage of the average width of the primary pie and secondary bar/pie.
     *
     * @param g
     */
    fun setpcGap(g: Int) {
        pcGap = g.toShort()
        val b = ByteTools.shortToLEBytes(pcGap)
        this.data[10] = b[0]
        this.data[11] = b[1]
    }

    /**
     * returns the distance between the primary pie and the secondary bar/pie.
     * The distance is specified as a percentage of the average width of the primary pie and secondary bar/pie.
     */
    fun getpcGap(): Int {
        return pcGap.toInt()
    }

    /**
     * specifies what determines the split between the primary pie and the secondary bar/pie via OOXML string value:
     * auto	-- split point of the chart group is determined automatically
     * pos --  The data is split based on the position of the data point in the series as specified by iSplitPos.
     * val --  The data is split based on a threshold value as specified by numSplitValue.
     * percent --The data is split based on a percentage threshold and the data point values represented as a percentage as specified by pcSplitPercent.
     * cust -- The data is split as arranged by the user.
     *
     * @param s
     */
    fun setSplitType(s: String?) {
        if (s == null) {
            fAutoSplit = true
            this.data[1] = 0
            return
        }
        fAutoSplit = false
        this.data[1] = 0
        if (s == "pos")
            split = 0
        else if (s == "val")
            split = 1
        else if (s == "percent")
            split = 2
        else if (s == "cust")
            split = 3
        val b = ByteTools.shortToLEBytes(split)
        this.data[2] = b[0]
        this.data[3] = b[1]
    }

    /**
     * Set specific options
     */
    override fun setChartOption(op: String, `val`: String): Boolean {
        if (op.equals("Gap", ignoreCase = true)) {
            setpcGap(Integer.parseInt(`val`))
        }
        return true
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 8071801452993935943L

        val prototype: XLSRecord?
            get() {
                val b = Boppop()
                b.opcode = XLSConstants.BOPPOP
                b.data = b.PROTOTYPE_BYTES
                b.init()
                return b
            }
    }
}
