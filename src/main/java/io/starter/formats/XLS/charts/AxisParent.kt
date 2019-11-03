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

import io.starter.formats.XLS.BiffRec
import io.starter.toolkit.ByteTools

/**
 * **AxisParent: Axis Size and Location (0x1014)**
 *
 *
 * This record specifies properties of an axis group and the beginning of a collection of records as defined by the chart sheet substream.
 *
 *
 *
 *
 * The Axis parent record stores most of the actual chart information,  what type of chart, x and y labels, etc.
 *
 *
 * 4	iax	2		axis index (0= main, 1= secondary)  This field MUST be set to zero when it is in the first AxisParent record in the chart sheet substream,
 * This field MUST be set to 1 when it is in the second AxisParent record in the chart sheet substream.
 * 16  	unused (16 bytes): Undefined and MUST be ignored.
 *
 *
 *
 *
 * this doesn't appear correct given ms doc
 * 6	x	4
 * 10	y	4
 * 14	dx	4		len of x axis
 * 18	dy	4		len of y axis
 */
class AxisParent : GenericChartObject(), ChartObject {
    private var iax: Short = 0
    private var x = 0
    private var y = 0
    private var dx = 0
    private var dy = 0

    /**
     * default version, returns the 1sst chart format record
     * <br></br>if there are more than 1 chart type in the chart
     * there will be multiple chart format records
     *
     * @return
     */
    protected val chartFormat: ChartFormat?
        get() = getChartFormat(0, false)

    /**
     * returns the plot area background color
     *
     * @return plot area background color hex string
     */
    //FormatHandle.COLOR_WHITE;
    val plotAreaBgColor: String?
        get() {
            for (i in chartArr.indices) {
                val b = chartArr[i]
                if (b is Frame) {
                    return b.bgColor
                }
            }
            return null
        }


    /**
     * returns true if this is a secondary axis
     *
     * @return
     */
    /**
     * sets if this is a secondary axis
     *
     * @param b
     */
    var isSecondaryAxis: Boolean
        get() = iax.toInt() == 1
        set(b) {
            if (b)
                iax = 1
            else
                iax = 0
            this.data[0] = iax.toByte()
        }

    override fun init() {
        super.init()
        iax = ByteTools.readShort(this.data!![0].toInt(), this.data!![1].toInt())
        x = ByteTools.readInt(this.getBytesAt(2, 4)!!)
        y = ByteTools.readInt(this.getBytesAt(6, 4)!!)
        dx = ByteTools.readInt(this.getBytesAt(10, 4)!!)
        dy = ByteTools.readInt(this.getBytesAt(14, 4)!!)

    }

    /**
     * get the chart format collection
     *
     * @return
     */
    protected fun getChartFormat(nChart: Int, addNew: Boolean): ChartFormat? {
        for (i in chartArr.indices) {
            val b = chartArr[i]
            if (b.opcode == XLSConstants.CHARTFORMAT) {
                val cf = b as ChartFormat
                if (cf.drawingOrder == nChart)
                    return cf
            }
        }
        return if (addNew) createNewChart(nChart) else null
    }

    /**
     * create basic cf necessary to define multiple or overlay charts
     * and add to this AxisParent
     *
     * @param nChart overlay chart number (>0 and <=9)
     * @return new ChartFormat
     */
    fun createNewChart(nChart: Int): ChartFormat? {
        // muliple charts, create one
        if (nChart > 0 && nChart <= 9) {
            val cf = ChartFormat.prototype as ChartFormat?
            cf!!.parentChart = this.parentChart
            val b = Bar.prototype as Bar?
            cf.chartArr.add(b)    // add a dummy chart object - will be replaced later
            val cfl = ChartFormatLink.prototype as ChartFormatLink?
            cf.chartArr.add(cfl)
            val sl = SeriesList()
            sl.opcode = XLSConstants.SERIESLIST
            sl.parentChart = this.parentChart
            cf.chartArr.add(sl)
            cf.drawingOrder = nChart
            this.chartArr.add(cf)    // add chartformat to chart array of axis parent
            return cf
        }
        return null
    }

    /**
     * remove axis records for pie-type charts ...
     */
    fun removeAxes() {
        if (chartArr[1] is Axis)
            chartArr.removeAt(1)    // remove Axis
        if (chartArr[1] is Axis)
            chartArr.removeAt(1)    // remove Axis
        if (chartArr[1] is TextDisp)
            chartArr.removeAt(1)    // remove Text for Axis
        if (chartArr[1] is TextDisp)
            chartArr.removeAt(1)    // remove Text for Axis
        if (chartArr[1] is PlotArea)
            chartArr.removeAt(1)    // remove
        if (chartArr[1] is Frame)
            chartArr.removeAt(1)    // remove Frame
        // all should be left is pos + chartFormat
    }

    /**
     * remove the desired axis + associated records
     *
     * @param axisType
     */
    fun removeAxis(axisType: Int) {
        // Remove axis
        for (i in chartArr.indices) {
            val b = chartArr[i]
            if (b.opcode == XLSConstants.AXIS) {
                if ((b as Axis).axis.toInt() == axisType) {
                    chartArr.removeAt(i)
                    break
                }
            }
        }
        val tdType = TextDisp.convertType(axisType)
        // Remove TextDisp assoc with Axis
        for (i in chartArr.indices) {
            val b = chartArr[i]
            if (b.opcode == XLSConstants.TEXTDISP) {
                val td = b as TextDisp
                if (tdType == td.type) {
                    chartArr.removeAt(i)
                    break
                }
            }
        }
    }

    /**
     * for those charts such as Radar, remove the
     * axis frame and plot area records as are not necessary
     */
    protected fun removePlotArea() {
        var remove = false
        var i = 0
        while (i < chartArr.size) {
            val b = chartArr[i]
            if (b.opcode == XLSConstants.AXIS) {
                remove = true
            } else if (b.opcode == XLSConstants.CHARTFORMAT) {
                break
            } else if (remove) {
                chartArr.removeAt(i)
                i--
            }
            i++
        }

    }

    /**
     * return XML of axis label and specific options for the desired axis
     * (both the Axis record + it's "associated" TextDisp record are consulted)
     *
     * @param Axis int desired Axis
     * @return String XML of axis label + specific options
     * @see ObjectLink
     */ // KSC: TODO: Get all axis ops ...
    fun getAxisOptionsXML(axis: Int): String {
        var bHasAxis = false
        var bHasLabel = false
        val sb = StringBuffer()
        val tdType = TextDisp.convertType(axis)
        for (i in chartArr.indices) {
            val b = chartArr[i]
            if (b.opcode == XLSConstants.AXIS) {
                if ((b as Axis).axis.toInt() == axis) {
                    sb.append(b.optionsXML)
                    bHasAxis = true
                }
            } else if (b.opcode == XLSConstants.TEXTDISP) {
                if ((b as TextDisp).type == tdType) {
                    sb.append(b.optionsXML)
                    bHasLabel = true
                    break
                }
            }
        }
        if (bHasAxis && !bHasLabel)
        // must include axis even if no textdisp/label
            sb.append(" Label=\"\"")
        return sb.toString()
    }

    /**
     * return the desired axis according to AxisType
     * Also finds and associates the axis rec with it's TextDisp
     * so as to be able to set axis optins ...
     * will create Axis + TextDisp if not present ...
     *
     * @param axisType
     * @param bCreateIfNecessary - if true, will create the axis if it doesn't exist
     * @return Desired axis
     */
    @JvmOverloads
    fun getAxis(axisType: Int, bCreateIfNecessary: Boolean = true): Axis? {
        var lastTd = 0
        var lastAxis = 0
        var a: Axis? = null
        var td: TextDisp? = null
        val tdType = TextDisp.convertType(axisType)
        var i = 0
        while (i < chartArr.size && td == null) {
            val b = chartArr[i]
            if (b.opcode == XLSConstants.AXIS) {
                lastAxis = i
                if ((b as Axis).axis.toInt() == axisType) {
                    a = b
                    a.setAP(this)    // ensure axis is linked to it's parent AxisParent 20090108 KSC:
                    //if (!bCreateIfNecessary) return a;
                }
            } else if (b.opcode == XLSConstants.TEXTDISP) {
                lastTd = i
                if (tdType == (b as TextDisp).type) {
                    td = b
                    if (bCreateIfNecessary) {// 20080723 KSC: added guard - but when is clearing td text necessary????
                        td.setText("")    // clear out axis legend
                    }
                }
            }
            i++
        }
        if (a == null && bCreateIfNecessary) {
            // if didn't find axis, then add axis + text disp ...
            // first, add TD
            td = TextDisp.getPrototype(axisType, "", this.wkbook) as TextDisp
            td.parentChart = this.parentChart
            this.chartArr.add(lastTd + 1, td)
            // next, add axis
            a = Axis.getPrototype(axisType) as Axis
            a.parentChart = this.parentChart
            this.chartArr.add(lastAxis + 1, a)
            a.setAP(this)    // ensure axis is linked to it's parent AxisParent 20090108 KSC:
        }
        if (a != null) {    // associate this axis with the textdisp if any
            a.td = td
        }
        return a
    }

    /**
     * sets the plot area background color
     *
     * @param bg color int
     */
    fun setPlotAreaBgColor(bg: Int) {
        for (i in chartArr.indices) {
            val b = chartArr[i]
            if (b is Frame) {
                b.setBgColor(bg)
                break
            }
        }
    }

    /**
     * adds a border around the plot area with the desired line width and line color
     *
     * @param lw
     * @param lclr
     */
    fun setPlotAreaBorder(lw: Int, lclr: Int) {
        for (i in chartArr.indices) {
            val b = chartArr[i]
            if (b is Frame) {
                b.addBox(lw, lclr, -1)
                break
            }
        }
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -2247258217367570732L
    }
}
/**
 * return the desired axis according to AxisType
 * (will create the axis if it doesn't exist)
 * Also finds and associates the axis rec with it's TextDisp
 * so as to be able to set axis optins ...
 * will create Axis + TextDisp if not present ...
 *
 * @param axisType
 * @return Desired axis
 */
