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

import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.XLS.Font
import io.starter.formats.cellformat.CellFormatFactory
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.json.JSONObject
import org.xmlpull.v1.XmlPullParser

import java.io.Serializable
import java.util.ArrayList
import java.util.HashMap
import java.util.Stack
import java.util.concurrent.atomic.AtomicInteger

/**
 * An axis group is a set of axes that specify a coordinate system, a set of chart groups that are plotted using these and the plot area that defines
 * where the axes are rendered on the chart.
 *
 *
 * In BIFF8, the AxisParent record governs the Axis Group.  A typical arrangement of records is:
 * AxisParent
 * Pos
 * Axis (X)
 * Axis (Y)
 * PlotArea
 * Frame
 * (Chart Group is one or more of:)
 * ChartFormat (defines the chart type) + Legend
 */
class ChartAxes
/**
 * The AxisParent record specifies if the axis group is the primary axis group or the secondary axis group on a chart.
 * Often the axes of the primary axis group are displayed to the left and bottom sides of the plot area, while axes of the secondary axis group are displayed on the right and top sides of the plot area.
 * The Pos record specifies the position and size of the outer plot area. The outer plot area is the bounding rectangle that includes the axis labels, the axis titles,
 * and data table of the chart. This record MUST be ignored on a secondary axis group.  */
/** The PlotArea record and the sequence of records that conforms to the FRAME rule in the sequence of records that conform to the AXES rule specify the properties of the inner plot area. The inner plot area is the rectangle bounded by the chart axes.
 * The PlotArea record MUST not exist on a secondary axis group.
 */
/**
 * @param ap
 */
(private var ap: AxisParent?) : ChartConstants, Serializable {
    internal var axes: ArrayList<*>? = ArrayList()
    internal var axisMetrics: HashMap<String, Any> = HashMap()    // holds important axis metrics such as xAxisReversed, xPattern, for outside use

    /**
     * returns the axis area background color hex string, or null if not set
     *
     * @return color hex string
     */
    val plotAreaBgColor: String?
        get() = ap!!.plotAreaBgColor

    /**
     * Returns true if the Y Axis (Value axis) is set to automatic scale
     *
     * The default setting for charts is known as Automatic Scaling
     * <br></br>When data changes, the chart automatically adjusts the scale (minimum, maximum values
     * plus major and minor tick units) as necessary
     *
     * @return boolean true if Automatic Scaling is turned on
     * @see getAxisAutomaticScale
     */
    /**
     * Sets the automatic scale option on or off for the Y Axis (Value axis)
     *
     * Automatic Scaling will automatically set the scale maximum, minimum and tick units
     * upon data changes, and is the default chart setting
     *
     * @param b
     * @see setAxisAutomaticScale
     */
    // group min and max together ...
    // TODO: throw exception?
    var axisAutomaticScale: Boolean
        get() {
            val a = getAxis(ChartConstants.YAXIS, false)
            return a?.isAutomaticScale ?: false
        }
        set(b) {
            val a = getAxis(ChartConstants.YAXIS, false)
            if (a != null)
                a.isAutomaticScale = b
        }


    /**
     * returns already-set axisMetrics
     * <br></br>Used only after calling chart.getMetrics
     *
     * @return axisMetrics -- map of useful chart display metrics
     * <br></br>Contains:
     * <br></br>
     * XAXISLABELOFFSET 	double
     * XAXISTITLEOFFSET	double
     * YAXISLABELOFFSET	double
     * YAXISTITLEOFFSET	double
     * xAxisRotate		integer x axis lable rotation angle
     * yPattern		string numeric pattern for y axis
     * xPattern		string numeric pattern for x axis
     * double major; // tick units
     * double minor;
     */
    val metrics: HashMap<*, *>
        get() = axisMetrics

    /**
     * store each axis
     *
     * @param a
     */
    fun add(a: Axis) {
        axes!!.add(a)
        a.setAP(this.ap)    // ensure axis is linked to it's parent AxisParent
    }

    fun setTd(axisType: Int, td: TextDisp) {
        val a = getAxis(axisType, false)
        if (a != null)
        // shouldn't
            a.td = td
    }

    /**
     * returns true if axisType is found on chart
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @return
     */
    fun hasAxis(axisType: Int): Boolean {
        for (i in axes!!.indices) {
            if ((axes!![i] as Axis).axis.toInt() == axisType)
                return true
        }
        return false
    }

    fun createAxis(axisType: Int) {
        val a = getAxis(axisType, true)
        return
    }

    /**
     * Return the desired axis
     * @return Axis
     *
     * public Axis getAxis(int axisType) {
     * return ap.getAxis(axisType);
     * }
     */

    /**
     * returns the desired axis if it exists
     * if bCreateIfNecessary, will create if it doesn't exist
     * otherwise returns null
     *
     * @param axisType           one of: XAXIS, YAXIS, ZAXIS
     * @param bCreateIfNecessary
     * @return
     */
    private fun getAxis(axisType: Int, bCreateIfNecessary: Boolean): Axis? {
        for (i in axes!!.indices) {
            if ((axes!![i] as Axis).axis.toInt() == axisType)
                return axes!![i] as Axis
        }
        if (bCreateIfNecessary) {
            axes!!.add(ap!!.getAxis(axisType, bCreateIfNecessary))
            return axes!![axes!!.size - 1] as Axis
        }
        return null
    }

    /**
     * returns true if Axis is reversed
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @return
     */
    protected fun isReversed(axisType: Int): Boolean {
        val a = getAxis(axisType, false)
        return a?.isReversed ?: false
    }

    /**
     * returns the number format for the desired axis
     *
     * @param axisType
     * @return number format string
     */
    protected fun getNumberFormat(axisType: Int): String? {
        val a = getAxis(axisType, false)
        return a?.numberFormat
    }

    /**
     * remove ALL axes from this chart
     */
    fun removeAxes() {
        ap!!.removeAxes()
        while (axes!!.size > 0) {
            axes!!.removeAt(0)
        }
    }

    /**
     * remove the desired axis + associated records
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     */
    fun removeAxis(axisType: Int) {
        ap!!.removeAxis(axisType)
        for (i in axes!!.indices) {
            if ((axes!![i] as Axis).axis.toInt() == axisType) {
                axes!!.removeAt(i)
                break
            }
        }
    }

    /**
     * return the coordinates of the outer plot area (plot + axes + labels) in pixels
     *
     * @return
     */
    fun getPlotAreaCoords(w: Float, h: Float): FloatArray? {
        val p = Chart.findRec(ap!!.chartArr, Pos::class.java) as Pos
        if (p != null) {
            val plotcoords = p.coords
            if (plotcoords != null) {
                plotcoords[0] = Pos.convertFromSPRC(plotcoords[0], w, 0f)
                plotcoords[1] = Pos.convertFromSPRC(plotcoords[1], 0f, h)
                plotcoords[2] = Pos.convertFromSPRC(plotcoords[2], w, 0f)    // SPRC units see Pos
                plotcoords[3] = Pos.convertFromSPRC(plotcoords[3], 0f, h)    // SPRC units see Pos
                return plotcoords
            }
        }
        return null
    }


    /**
     * sets the plot area background color
     *
     * @param bg color int
     */
    fun setPlotAreaBgColor(bg: Int) {
        ap!!.setPlotAreaBgColor(bg)
    }

    /**
     * adds a border around the plot area with the desired line width and line color
     *
     * @param lw
     * @param lclr
     */
    fun setPlotAreaBorder(lw: Int, lclr: Int) {
        ap!!.setPlotAreaBorder(lw, lclr)
    }


    /**
     * obtain the desired axis' label and other options, if present, in XML form
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @return String XML representation of desired axis' label and options
     * @see ObjectLink
     */
    fun getAxisOptionsXML(axisType: Int): String {
        return ap!!.getAxisOptionsXML(axisType)
    }

    /**
     * Return the Axis Title, or "" if none
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     */
    fun getTitle(axisType: Int): String {
        val a = getAxis(axisType, false)
        return if (a != null) a.title else ""
    }

    /**
     * set the axis title string
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @param Title
     */
    fun setTitle(axisType: Int, Title: String) {
        val a = getAxis(axisType, false)
        if (a != null) {
            val defaultIsVert = axisType == ChartConstants.YAXIS && "" == a.title
            a.title = Title
            if (defaultIsVert) a.td!!.rotation = 90
        }
    }

    /**
     * return the rotation of the axis labels, if any
     * 0 to 90		Text rotated 0 to 90 degrees counter-clockwise
     * 91 to 180	Text rotated 1 to 90 degrees clockwise (angle is trot – 90)
     * 255			Text top-to-bottom with letters upright
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     */
    fun getLabelRotation(axisType: Int): Int {
        var rot = 0
        val a = getAxis(axisType, false)
        if (a != null) {
            val t = Chart.findRec(a.chartArr, Tick::class.java) as Tick
            if (t != null) {    // shoudn't
                rot = t.rotation.toInt()
                when (rot) {
                    0 -> {
                    }
                    1        // not correct ...
                    -> rot = 180
                    2 -> rot = -90
                    3 -> rot = 90
                }
            }
        }
        return rot
    }

    /**
     * return the rotation of the axis Title, if any
     * 0 to 90		Text rotated 0 to 90 degrees counter-clockwise
     * 91 to 180	Text rotated 1 to 90 degrees clockwise (angle is trot – 90)
     * 255			Text top-to-bottom with letters upright
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     */
    fun getTitleRotation(axisType: Int): Int {
        val a = getAxis(axisType, false)
        if (a != null) {
            val td = a.td
            return td!!.rotation
        }
        return 0
    }

    /**
     * return the coordinates, in pixels, of the title text area, if possible
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @return float[]
     */
    fun getCoords(axisType: Int): FloatArray? {
        val a = getAxis(axisType, false)
        if (a != null) {
            val td = a.td
            return td!!.coords
        }
        return floatArrayOf(0f, 0f, 0f, 0f)
    }

    /**
     * Returns the scale values of the the desired Value axis
     * <br></br>Scale elements: (min value, max value, minor (tick) scale, major (tick) scale)
     *
     * The Value axis contains numbers rather than labels, and is normally the Y axis,
     * but Scatter and Bubble charts may have a value axis on the X Axis as well
     *
     * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS, ChartHandle.ZAXIS
     * @return int Miminum Scale value for the desired axis
     * @see getAxisMinScale
     */
    @JvmOverloads
    fun getMinMax(ymin: Double, ymax: Double, axisType: Int = ChartConstants.YAXIS): DoubleArray {
        val a = getAxis(axisType, false)
        val ret = DoubleArray(4)
        if (a != null) {
            val v = Chart.findRec(a.chartRecords, ValueRange::class.java) as ValueRange
            if (v.isAutomaticScale)
            // major crazy Excel "automatic max/min/tickmarks calc" ... a monster
                v.setMaxMin(ymax, ymin)
            ret[0] = v.min
            ret[1] = v.max
            ret[2] = v.minorTick
            ret[3] = v.majorTick
        }
        return ret
    }

    /**
     * Returns the major tick unit of the desired Value axis
     *
     * The Value axis contains numbers rather than labels, and is normally the Y axis,
     * but Scatter and Bubble charts may have a value axis on the X Axis as well
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @return int major tick unit
     * @see getAxisMajorUnit
     */
    fun getAxisMajorUnit(axisType: Int): Int {
        val a = getAxis(axisType, false)
        return a?.majorUnit ?: 10
// TODO throw exception instead ???
    }

    /**
     * Returns the minor tick unit of the desired Value axis
     *
     * The Value axis contains numbers rather than labels, and is normally the Y axis,
     * but Scatter and Bubble charts may have a value axis on the X Axis as well
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @return int - minor tick unit of the desired axis
     * @see getAxisMinorUnit
     */
    fun getAxisMinorUnit(axisType: Int): Int {
        val a = getAxis(axisType, false)
        return a?.minorUnit ?: 1
// TODO throw exception instead
    }

    /**
     * Sets the maximum scale value of the desired Value axis
     *
     * The Value axis contains numbers rather than labels, and is normally the Y axis,
     * but Scatter and Bubble charts may have a value axis on the X Axis as well
     *
     * Note: The default scale setting for charts is known as Automatic Scaling
     * <br></br>When data changes, the chart automatically adjusts the scale as necessary
     * <br></br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @param int      MaxValue - desired maximum value of the desired axis
     * @see setAxisMax
     */
    fun setAxisMax(axisType: Int, MaxValue: Int) {
        val a = getAxis(axisType, false)
        a?.setMaxScale(MaxValue)
        // TODO: throw exception if no axis?
    }

    /**
     * Sets the minimum scale value of the desired Value axis
     *
     * The Value axis contains numbers rather than labels, and is normally the Y axis,
     * but Scatter and Bubble charts may have a value axis on the X Axis as well
     *
     * Note: The default setting for charts is known as Automatic Scaling
     * <br></br>When data values change, the chart automatically adjusts the scale as necessary
     * <br></br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @param int      MinValue - the desired Minimum scale value
     * @see setAxisMin
     */
    fun setAxisMin(axisType: Int, MinValue: Int) {
        val a = getAxis(axisType, false)
        a?.setMinScale(MinValue)
        // TODO: throw exception if no axis?
    }

    /**
     * Sets the automatic scale option on or off for the desired Value axis
     *
     * The Value axis contains numbers rather than labels, and is normally the Y axis,
     * but Scatter and Bubble charts may have a value axis on the X Axis as well
     *
     * Automatic Scaling automatically sets the scale maximum, minimum and tick units
     * upon data changes, and is the default setting for charts
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @param boolean  b - true if set Automatic scaling on, false otherwise
     * @see setAxisAutomaticScale
     */
    fun setAxisAutomaticScale(axisType: Int, b: Boolean) {
        val a = getAxis(axisType, false)
        if (a != null)
            a.isAutomaticScale = b
    }

    /**
     * Returns true if the desired Value axis is set to automatic scale
     *
     * The Value axis contains numbers rather than labels, and is normally the Y axis,
     * but Scatter and Bubble charts may have a value axis on the X Axis as well
     *
     * Note: The default setting for charts is known as Automatic Scaling
     * <br></br>When data changes, the chart automatically adjusts the scale as necessary
     * <br></br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @return boolean true if Automatic Scaling is turned on
     * @see getAxisAutomaticScale
     */
    fun getAxisAutomaticScale(axisType: Int): Boolean {
        val a = getAxis(axisType, false)
        return a?.isAutomaticScale ?: false    // group min and max together ...
// TODO: throw exception?
    }

    /**
     * Sets the maximum value of the Y Axis (Value Axis) Scale
     *
     * Note: The default scale setting for charts is known as Automatic Scaling
     * <br></br>When data changes, the chart automatically adjusts the scale as necessary
     * <br></br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
     *
     * @param int MaxValue - the desired maximum scale value
     * @see ChartHandle.setAxisMax
     */
    fun setAxisMax(MaxValue: Int) {
        val a = getAxis(ChartConstants.YAXIS, false)
        a?.setMaxScale(MaxValue)
        // TODO: throw exception if no axis?
    }

    /**
     * Sets the minimum value of the Y Axis (Value Axis) Scale
     *
     * Note: The default setting for charts is known as Automatic Scaling
     * <br></br>When data changes, the chart automatically adjusts the scale as necessary
     * <br></br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
     *
     * @param int MinValue - the desired minimum scale value
     * @see ChartHandle.setAxisMin
     */
    fun setAxisMin(MinValue: Int) {
        val a = getAxis(ChartConstants.YAXIS, false)
        a?.setMinScale(MinValue)
        // TODO: throw exception if no axis?
    }

    /**
     * returns the SVG necesssary to define the desired axis
     *
     * @param axisType     one of: XAXIS, YAXIS, ZAXIS
     * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
     * @return
     */
    fun getSVG(axisType: Int, chartMetrics: Map<String, Double>, categories: Array<Any>): String? {
        val a = getAxis(axisType, false)
        return a?.getSVG(this, chartMetrics, categories)
    }

    /**
     * returns the OOXML necessary to define the desired axis
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @param type
     * @param id
     * @param crossId
     * @return
     */
    fun getOOXML(axisType: Int, type: Int, id: String, crossId: String): String {
        val a = getAxis(axisType, false)
        return if (a != null) a.getOOXML(type, id, crossId) else ""
    }

    /**
     * set the font index for this Axis (for title)
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @param fondId
     */
    fun setTitleFont(axisType: Int, fondId: Int) {
        val a = getAxis(axisType, false)
        a?.setFont(fondId)
    }

    /**
     * returns the font for the title for the desired axis
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @return
     */
    fun getTitleFont(axisType: Int): Font? {
        val a = getAxis(axisType, false)
        return a?.font
    }

    /**
     * returns the label font for the desired axis.
     * If not explicitly set, returns the default font
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @return
     */
    fun getLabelFont(axisType: Int): Font? {
        val a = getAxis(axisType, false)
        return a?.labelFont
    }


    fun getJSON(axisType: Int, wbh: io.starter.OpenXLS.WorkBookHandle, chartType: Int, yMax: Double, yMin: Double, nSeries: Int): JSONObject {
        return ap!!.getAxis(axisType, false)!!.getJSON(wbh, chartType, yMax, yMin, nSeries)
    }


    /**
     * sets an option for this axis
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @param op       option name; one of: CatCross, LabelCross, Marks, CrossBetween, CrossMax,
     * MajorGridLines, AddArea, AreaFg, AreaBg
     * or Linked Text Display options:
     * Label, ShowKey, ShowValue, ShowLabelPct, ShowPct,
     * ShowCatLabel, ShowBubbleSizes, TextRotation, Font
     * @param val      option value
     */
    fun setChartOption(axisType: Int, op: String, `val`: String) {
        val a = getAxis(axisType, false)
        a?.setChartOption(op, `val`)
    }


    /**
     * returns the Axis Label Placement or position as an int
     *
     * One of:
     * <br></br>Axis.INVISIBLE - axis is hidden
     * <br></br>Axis.LOW - low end of plot area
     * <br></br>Axis.HIGH - high end of plot area
     * <br></br>Axis.NEXTTO- next to axis (default)
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @return int - one of the Axis Label placement constants above
     */
    fun getAxisPlacement(axisType: Int): Int {
        val a = getAxis(axisType, false)
        return a?.axisPlacement ?: -1
    }

    /**
     * sets the axis labels position or placement to the desired value (these match Excel placement options)
     *
     * Possible options:
     * <br></br>Axis.INVISIBLE - hides the axis
     * <br></br>Axis.LOW - low end of plot area
     * <br></br>Axis.HIGH - high end of plot area
     * <br></br>Axis.NEXTTO- next to axis (default)
     *
     * @param axisType  one of: XAXIS, YAXIS, ZAXIS
     * @param Placement - int one of the Axis placement constants listed above
     */
    fun setAxisPlacement(axisType: Int, placement: Int) {
        val a = getAxis(axisType, false)
        if (a != null)
            a.axisPlacement = placement
    }


    /**
     * parse OOXML axis element
     *
     * @param axisType one of: XAXIS, YAXIS, ZAXIS
     * @param xpp      XmlPullParser positioned at correct elemnt
     * @param axisTag  catAx, valAx, serAx, dateAx
     * @param lastTag  Stack of element names
     */
    fun parseOOXML(axisType: Int, xpp: XmlPullParser, tnm: String, lastTag: Stack<String>, bk: WorkBookHandle) {
        val a = getAxis(axisType, true)
        if (a != null) {
            a.removeTitle()
            a.setChartOption("MajorGridLines", "false")    // initally, remove any - will be set in parseOOXML if necessary
            a.parseOOXML(xpp, tnm, lastTag, bk)
        }
    }

    fun close() {
        axes!!.clear()
        axes = null
        ap = null
        axisMetrics.clear()
    }

    /**
     * returns a specific axis Metric
     * <br></br>Use only after calling chart.getMetrics
     *
     * @param metric String metric option one of
     * <br></br>
     * XAXISLABELOFFSET 	double
     * XAXISTITLEOFFSET	double
     * YAXISLABELOFFSET	double
     * YAXISTITLEOFFSET	double
     * xAxisRotate		integer x axis lable rotation angle
     * yPattern		string numeric pattern for y axis
     * xPattern		string numeric pattern for x axis
     * double major; // tick units
     * double minor;
     * @return
     */
    fun getMetric(metric: String): Any {
        return axisMetrics[metric]
    }

    /**
     * generate Axis Metrics -- minor, major scale, title and label offsets, etc.
     * for use in SVG generation and other operations
     *
     * @param charttype    chart type constant
     * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
     * @param plotcoords   == if set, plot-area coordinates (may not be set if automatic)
     * @return axisMetrics Map:  contains the important axis metrics for use in chart display
     */
    fun getMetrics(charttype: Int, chartMetrics: HashMap<String, Double>, plotcoords: FloatArray, categories: Array<Any>): HashMap<String, Any> {
        val minmax = this.getMinMax(chartMetrics["min"], chartMetrics["max"])    // sets min/max on Value axis, based upon axis settings and actual minimun and maximum values
        chartMetrics["min"] = minmax[0]    // set new values, if any
        chartMetrics["max"] = minmax[1]    // ""
        axisMetrics["minor"] = minmax[2]
        axisMetrics["major"] = minmax[3]
        axisMetrics["xAxisReversed"] = this.isReversed(ChartConstants.XAXIS) // default is Value axis crosses at bottom. reverse= crosses at top
        axisMetrics["xPattern"] = this.getNumberFormat(ChartConstants.XAXIS)
        axisMetrics["yAxisReversed"] = this.isReversed(ChartConstants.YAXIS) // if value (Y), default is on LHS; reverse= RHS
        axisMetrics["yPattern"] = this.getNumberFormat(ChartConstants.YAXIS)
        axisMetrics["XAXISLABELOFFSET"] = 0.0
        axisMetrics["XAXISTITLEOFFSET"] = 0.0
        axisMetrics["YAXISLABELOFFSET"] = 0.0
        axisMetrics["YAXISTITLEOFFSET"] = 0.0

        // X axis title Offset
        if (this.getTitle(ChartConstants.XAXIS) != "") {
            val ef = this.getTitleFont(ChartConstants.XAXIS)
            val f = java.awt.Font(ef!!.fontName, ef.fontWeight, ef.fontHeightInPoints.toInt())
            val h = AtomicInteger(0)
            val w = getRotatedWidth(f, h, this.getTitle(ChartConstants.XAXIS), this.getTitleRotation(ChartConstants.XAXIS))
            axisMetrics["XAXISTITLEOFFSET"] = h.toInt() + 10    /* add a little padding */
        }

        // Y Axis Title Offsets
        if (this.hasAxis(ChartConstants.YAXIS)) {
            val title = this.getTitle(ChartConstants.YAXIS)
            if (title != "") {
                val ef = this.getTitleFont(ChartConstants.YAXIS)
                val f = java.awt.Font(ef!!.fontName, ef.fontWeight, ef.fontHeightInPoints.toInt())
                val h = AtomicInteger(0)
                val w = getRotatedWidthVert(f, h, this.getTitle(ChartConstants.YAXIS), this.getTitleRotation(ChartConstants.YAXIS))
                axisMetrics["YAXISTITLEOFFSET"] = w // add padding
                /*
    			int rot= this.getTitleRotation(YAXIS); //0-180 or 255 (=vertical with letters upright)
				if (rot==90 || rot==255 || charttype==BARCHART)	{// Y axis title is almost always rotated, so use font height as offset
					if (charttype!=BARCHART)
						axisMetrics.put("YAXISTITLEOFFSET", Math.ceil(ef.getFontHeightInPoints()*xpaddingfactor));
					else
						axisMetrics.put("XAXISTITLEOFFSET", Math.ceil(ef.getFontHeightInPoints()*ypaddingfactortitle));
				} else {// get width in desired rotation ****
					if (rot > 180) rot-=180;	// ensure angle is 0-90
	    			java.awt.Font f= new java.awt.Font(ef.getFontName(), ef.getFontWeight(), (int)ef.getFontHeightInPoints());
		    		double len= StringTool.getApproximateStringWidth(f, ExcelTools.getFormattedStringVal(title, null));
		    		axisMetrics.put("YAXISTITLEOFFSET", Math.ceil(len*(Math.cos(Math.toRadians(rot)))));
				}*/
            } else
                axisMetrics["YAXISTITLEOFFSET"] = 20.0    // no axis title- add a little padding from edge
        }
        // Label offsets
        var series = arrayOfNulls<StringBuffer>(0)
        try {
            val major = getMetric("major") as Double
            val min = chartMetrics["min"]
            val max = chartMetrics["max"]
            // get java awt font so can compute approximate width of largest y axis label in order to compute YAXISLABELOFFSET
            if (major > 0) {
                var nSeries = 0
                if (min == 0.0)
                // usual case
                    nSeries = (if (major != 0.0) (max / major).toInt() + 1 else 0)
                else
                    nSeries = if (major != 0.0) ((max - min) / major).toInt() + 1 else 0
                nSeries = Math.abs(nSeries)
                series = arrayOfNulls(nSeries)
                if (Math.floor(major) != major) {    // contains a fractional part ... avoid java floating point issues
                    // ensure y value matches scale/avoid double precision issues ... sigh ...
                    val s = major.toString()
                    var z = s.indexOf(".")
                    var scale = 0
                    if (z != -1)
                        scale = s.length - (z + 1)
                    z = 0
                    var i = min
                    while (i <= max) {
                        val bd = java.math.BigDecimal(i).setScale(scale, java.math.BigDecimal.ROUND_HALF_UP)
                        series[z++] = StringBuffer(
                                CellFormatFactory.fromPatternString(
                                        getMetric("yPattern") as String).format(bd.toString()))
                        i += major
                    }
                } else { // usual case of an int scale
                    var z = 0
                    var i = min
                    while (i <= max) {
                        series[z++] = StringBuffer(CellFormatFactory.fromPatternString(getMetric("yPattern") as String).format(i))
                        i += major
                    }
                }
            }
        } catch (e: Exception) {
            Logger.logWarn("ChartAxes.getMetrics.  Error obtaining Series: $e")
        }

        // Label Offsets ...
        if (this.hasAxis(ChartConstants.XAXIS) && charttype != ChartConstants.RADARCHART && charttype != ChartConstants.RADARAREACHART) { //(Pie, donut, etc. don't have axes labels so disregard
            val s: Array<Any>            // Determine X Axis Label offsets
            val width: Double
            val rot: AtomicInteger
            val lf: io.starter.formats.XLS.Font?
            val pattern: String
            if (charttype != ChartConstants.BARCHART) {
                s = categories
                pattern = getMetric("xPattern") as String
            } else {        // bar chart - series on x axis
                s = series
                pattern = getMetric("yPattern") as String
            }
            width = chartMetrics["w"] / s.size - 6 // ensure a bit of padding on either side
            lf = this.getLabelFont(ChartConstants.XAXIS)
            rot = AtomicInteger(this.getLabelRotation(ChartConstants.XAXIS))        // if rot==0 and xaxis labels do not fit in width, a forced rotation will happen. ...
            val off = getLabelOffsets(lf, width, s, rot, pattern, true)
            axisMetrics["xAxisRotate"] = rot.toInt() // possibly changed when calculating label offsets
            axisMetrics["XAXISLABELOFFSET"] = off
        }
        if (this.hasAxis(ChartConstants.YAXIS) && charttype != ChartConstants.RADARCHART && charttype != ChartConstants.RADARAREACHART) {    //(Pie, Donut, etc. don't have axes labels so disregard
            // for Y axis, determine width of labels and use as offset (except for bar charts, use height as offset)
            val s: Array<Any>
            val width: Double
            val rot: AtomicInteger
            val lf: io.starter.formats.XLS.Font?
            val pattern: String
            if (charttype != ChartConstants.BARCHART) {
                s = series
                pattern = getMetric("yPattern") as String
            } else {    // bar chart- categories on y axis
                s = categories
                pattern = getMetric("xPattern") as String
            }
            rot = AtomicInteger(this.getLabelRotation(ChartConstants.YAXIS))        // if rot==0 and xaxis labels do not fit in width, a forced rotation will happen. ...
            lf = this.getLabelFont(ChartConstants.YAXIS)
            width = chartMetrics["w"] / 2 - 10 // a good guess?
            val off = getLabelOffsets(lf, width, s, rot, pattern, false)
            axisMetrics["yAxisRotate"] = rot.toInt()    // possibly changed when calculating label offsets
            axisMetrics["YAXISLABELOFFSET"] = off // with padding
        }
        return axisMetrics
    }

    /**
     * returns the offset of the labels for the axis given the length of the label strings, the width of the
     * <br></br>this will attempt to  break apart longer strings; if so will increment offset to accomodate multiple lines
     *
     * @param lf      label font
     * @param width   max width of labels
     * @param strings label strings
     * @param rot     desired rotation, if any (0 if none)
     * @param horiz   true if this is a horizontal axis
     * @return
     */
    private fun getLabelOffsets(lf: io.starter.formats.XLS.Font?, width: Double, strings: Array<Any>, rot: AtomicInteger, pattern: String, horiz: Boolean): Double {
        var width = width
        var retwidth = 0.0
        var f: java.awt.Font? = null
        var h = 0.0
        try {
            // get awt Font so can compute and fit category in width
            f = java.awt.Font(lf!!.fontName, lf.fontWeight, lf.fontHeightInPoints.toInt())
            for (i in strings.indices) {
                var w: Double
                val height = AtomicInteger(0)
                var s: StringBuffer? = null
                try {
                    s = StringBuffer(CellFormatFactory.fromPatternString(pattern).format(strings[i].toString()))
                } catch (e: IllegalArgumentException) { // trap error formatting
                    s = StringBuffer(strings[i].toString())
                }

                if (horiz)
                // on horizontal axis
                    w = getRotatedWidth(f, height, s!!.toString(), rot.toInt())
                else
                // on vertical axis
                    w = getRotatedWidthVert(f, height, s!!.toString(), rot.toInt())
                h = Math.max(height.toInt().toDouble(), h)
                w = addLinesToFit(f, s, rot, w, width, height)
                if (w > width)
                    width = w

                strings[i] = s
                retwidth = Math.max(w, retwidth)
            }
        } catch (e: Exception) {
        }

        return if (horiz)
            h
        else
            retwidth
        //		return offset + h;
    }

    /**
     * get the width of string s in font f rotated by rot (0, 90, -90, 180)
     *
     * @param f   Font to display s in
     * @param s   string to display
     * @param rot rotation (0= none)
     */
    private fun getRotatedWidth(f: java.awt.Font, height: AtomicInteger, s: String, rot: Int): Double {
        var retWidth = 0.0
        val slines = s.split("\n".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
        for (i in slines.indices) {
            var width: Double
            if (rot == 0) {
                width = StringTool.getApproximateStringWidth(f, CellFormatFactory.fromPatternString(null).format(slines[i]))
                height.set(Math.ceil((f.size * 3).toDouble()).toInt()) // width of the font + padding
            } else if (Math.abs(rot) == 90) {
                width = Math.ceil((f.size * 2).toDouble()) // width of the font + padding
                height.set(Math.max(height.toInt().toDouble(), StringTool.getApproximateStringWidth(f, CellFormatFactory.fromPatternString(null).format(slines[i]))).toInt())
            } else { // 45
                width = StringTool.getApproximateStringWidth(f, CellFormatFactory.fromPatternString(null).format(slines[i]))
                width = Math.ceil(width * Math.cos(Math.toRadians(rot.toDouble())))
                height.set(Math.ceil(width * Math.sin(Math.toRadians(rot.toDouble()))).toInt())
            }
            retWidth = Math.max(width, retWidth)
        }
        return retWidth
    }

    /**
     * get the width of string s in font f rotated by rot (0, 90, -90, 180)
     * on a Vertical axis
     *
     * @param f   Font to display s in
     * @param s   string to display
     * @param rot rotation (0= none)
     */
    private fun getRotatedWidthVert(f: java.awt.Font, height: AtomicInteger, s: String, rot: Int): Double {
        var retWidth = 0.0
        val slines = s.split("\n".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
        for (i in slines.indices) {
            var width: Double
            if (Math.abs(rot) == 90) {    // means VERTICAL orientation - default
                height.set(StringTool.getApproximateStringWidth(f, CellFormatFactory.fromPatternString(null).format(slines[i])).toInt())
                width = Math.ceil((f.size * 2).toDouble()) // width of the font + padding
            } else if (rot == 0) {    // means HORIZONTAL on vertical axis
                height.set(Math.ceil((f.size * 3).toDouble()).toInt()) // width of font
                width = Math.max(height.toInt().toDouble(), StringTool.getApproximateStringWidth(f, CellFormatFactory.fromPatternString(null).format(slines[i])))
            } else { // 45
                width = StringTool.getApproximateStringWidth(f, CellFormatFactory.fromPatternString(null).format(slines[i]))
                width = Math.ceil(width * Math.cos(Math.toRadians(rot.toDouble())))
                height.set(Math.ceil(width * Math.sin(Math.toRadians(rot.toDouble()))).toInt())
            }
            retWidth = Math.max(width, retWidth)
        }
        return retWidth
    }

    /**
     * if formatted width of string doesn't fit into width, break apart at spaces with new lines
     * return
     *
     * @param f     font to display string in
     * @param s     target String
     * @param rot   rotation (0, 45, 90, -90 or 180)
     * @param len   formatted string length taking into account display font
     * @param width maximum width to display string
     * @return
     */
    private fun addLinesToFit(f: java.awt.Font, s: StringBuffer, rot: AtomicInteger, len: Double, width: Double, height: AtomicInteger): Double {
        var len = len
        var retLen = Math.min(width, len)
        var str = s.toString().trim { it <= ' ' }

        while (len > width) {
            var lastSpace = -1
            var j = s.lastIndexOf("\n") + 1
            len = -1.0
            while (len < width && j < str.length) {
                len += StringTool.getApproximateCharWidth(f, str.get(j))
                if (str.get(j) == ' ')
                    lastSpace = j
                j++
            }
            if (len < width)
                break    // got it

            if (lastSpace == -1) {    // no spaces to break apart via \n's - rotate!
                if (str.indexOf(' ') == -1) {
                    // see if string will fit in 45 degree rotation
                    if (rot.toInt() != -90) {
                        rot.set(45)
                        len = getRotatedWidth(f, height, str, rot.toInt())
                        if (len > width) {    // then all's fine
                        } else
                        // doesn't fit in 45, must be 90 degrees
                            rot.set(-90)
                    }
                    if (rot.toInt() == -90) {
                        len = getRotatedWidth(f, height, str, rot.toInt())
                    }
                    retLen = Math.max(len, retLen)
                    break
                } else
                    lastSpace = s.toString().indexOf(' ')
            }
            s.replace(lastSpace, lastSpace + 1, "\n")    // + str.substring(lastSpace+1));
            str = s.toString()
        }
        return retLen
    }

    companion object {
        /**
         * serialVersionUID
         */
        private const val serialVersionUID = -7862828186455339066L
    }
}
/**
 * Returns the YAXIS scale elements (min value, max value, minor (tick) scale, major (tick) scale)
 * the scale elements are calculated from the minimum and maximum values on the chart
 *
 * @param ymin the minimum value of all the series values
 * @param ymax the maximum value of all the series values
 * @return double[] min, max, minor, major
 */
