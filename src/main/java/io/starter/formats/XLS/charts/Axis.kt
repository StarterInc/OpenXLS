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

import io.starter.OpenXLS.CellRange
import io.starter.OpenXLS.ChartHandle
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.OOXML.*
import io.starter.formats.XLS.BiffRec
import io.starter.formats.XLS.Font
import io.starter.formats.XLS.XLSRecord
import io.starter.formats.cellformat.CellFormatFactory
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.json.JSONArray
import org.json.JSONException
import org.json.JSONObject
import org.xmlpull.v1.XmlPullParser

import java.util.Stack

//OOXML-specific structures
//OOXML-specific structures

/**
 * **Axis: Axis Type (0x101d)**
 *
 *
 * 4	wType		2	axis type (0= category axis or x axis on a scatter chart, 1= value axis, 2= series axis
 * 6	(reserved)	16	0
 *
 *
 * Order of Axis Subrecords:
 * X (Cat Axis)
 *
 *
 * CatSerRange
 * AxcExt
 * CatLab
 * [IfmtRecord]
 * [Tick]
 * [FontX]
 * [AxisLineFormat, LineFormat] - if gridlines (major or minor) or border around wall/floor
 * [AreaFormat, GelFrame]
 * [Shape or TextPropsStream]
 * [CrtMltFrt]
 *
 *
 * Y (Value Axis)
 * ValueRange
 * [YMult -->Text-->Pos...] -- display units
 * [IfmtRecord]
 * [Tick]
 * [FontX]
 * [AxisLineFormat, LineFormat] - if gridlines (major or minor) or border around wall/floor
 * [AreaFormat, GelFrame]
 * [Shape or TextPropsStream]
 * [CrtMltFrt]
 *
 *
 * XY (Series Axis)
 * CatSerRange
 * [IfmtRecord]
 * [Tick]
 * [FontX]
 * [AxisLineFormat, LineFormat] - if gridlines (major or minor) or border around wall/floor
 * [AreaFormat, GelFrame]
 * [Shape or TextPropsStream]
 * [CrtMltFrt]
 */
/* General info on Excel axis types:
 *  In Microsoft Excel charts, there are different types of X axes.
 *  While the Y axis is a Value type axis (i.e. containing a ValueRange record),
 *  the X axis can be a Category type axis or a Value type axis.
 *  Using a Value axis, the data is treated as continuously varying numerical data,
 *  and the marker is placed at a point along the axis which varies according to its
 *  numerical value.
 *  Using a Category axis, the data is treated as a sequence of non-numerical text labels,
 *  and the marker is placed at a point along the axis according to its position in
 *  the sequence.
 *
 *  Note that Scatter (x/y) charts and Bubble charts have two Value axes, pie and donut charts have no axes
 *
 *  How do you arrange your chart so the categories are displayed along the Y axis?
 *  The method involves adding a dummy series along the Y axis,
 *  applying data labels to its points for category labels,
 *  and making the original Y axis disappear.
 */
class Axis : GenericChartObject(), ChartObject {
    var axis: Short = 0
        internal set
    private var linkedtd: TextDisp? = null    // 20070730 KSC: hold axis legend TextDisp for this axis
    private var ap: AxisParent? = null        // 20090108 KSC: links to this axes parent --
    // OOXML-specific
    /**
     * return the OOXML shape property for this axis
     *
     * @return
     */
    /**
     * define the OOXML shape property for this axis from an existing spPr element
     */
    //shapeProps.setNS("c");
    var spPr: SpPr? = null    // 20081224 KSC:  OOXML-specific holds the shape properties (line and fill) for this axis
    /**
     * return the OOXML title element for this axis
     *
     * @return
     */
    /**
     * set the OOXML title element for this axis
     *
     * @param t
     */
    var ooxmlTitle: io.starter.formats.OOXML.Title? = null        // OOXML title element
    private var txpr: TxPr? = null        // text properties for axis
    private var nf: NumFmt? = null        // NumFmt prop for axis
    internal var axPos: String? = null

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    /**
     * @return String XML representation of this chart-type's options
     */
    override// Axis 0: CatserRange, Axcent, Tick [AxisLineFormat, LineFormat, AreaFormat] 					last 3 recs are for 3d formatting
    // Axis 1: ValueRange, Tick, AxisLineFormat, LineFormat [AreaFormat, LineFormat, AreaFormat]		"	"
    // Axis 2: [CatserRange, Tick]		Z axis, for surface charts (only??)
    // Handle subordinate record options here rather than in the specific rec
    // record deviations from defaults
    // necessary for 3d charts and bg colors ...
    // Y Axis draw gridline usually true, trap if do not have (see below)
    // indicates area
    // should ONLY be present if AxisLineFormat AddArea
    // if not default background colors, record and reset
    // if not defaults, set fore and back of walls, sides or floor
    // xaxis
    // yaxis
    // xaxis
    // yaxis
    // TODO: parse LineFormat and, if not standard, add as option
    // TODO: Parse Tick for options
    // most Y Axes have major grid lines; flag if do not
    val optionsXML: String
        get() {
            val sb = StringBuffer()
            var hasMajorGridLines = false
            for (i in chartArr.indices) {
                val b = chartArr[i]
                if (b is CatserRange) {
                    if (b.catCross != 1)
                        sb.append(" CatCross=\"" + b.catCross + "\"")
                    if (b.catLabel != 1)
                        sb.append(" LabelCross=\"" + b.catLabel + "\"")
                    if (b.catMark != 1)
                        sb.append(" Marks=\"" + b.catMark + "\"")
                    if (!b.crossBetween)
                        sb.append(" CrossBetween=\"false\"")
                    if (b.crossMax)
                        sb.append(" CrossMax=\"true\"")
                } else if (b is AxisLineFormat) {
                    val id = b.id.toInt()
                    if (id == AxisLineFormat.ID_MAJOR_GRID)
                        hasMajorGridLines = true
                    else if (id == AxisLineFormat.ID_WALLORFLOOR)
                        sb.append(" AddArea=\"true\"")
                } else if (b is AreaFormat) {
                    var icvFore = b.geticvFore()
                    var icvBack = b.geticvBack()
                    if (icvBack == -1) icvBack = 0x4D
                    if (icvFore == 1) icvFore = 0x4E
                    if (axis.toInt() == 0 && icvFore != 22 || axis.toInt() == 1 && icvFore != -1)
                        sb.append(" AreaFg=\"$icvFore\"")
                    if (axis.toInt() == 0 && icvBack != 0 || axis.toInt() == 1 && icvBack != 1)
                        sb.append(" AreaBg=\"$icvBack\"")
                }
            }
            if (axis.toInt() == 1 && !hasMajorGridLines)
                sb.append(" MajorGridLines=\"false\"")
            return sb.toString()
        }

    /**
     * get/set linked TextDisp (legend for this axis)
     * used for setting axis options
     *
     * @param td
     */
    var td: TextDisp?
        get() {
            if (linkedtd == null)
                getTitleTD(false)
            return linkedtd
        }
        set(td) {
            linkedtd = td
        }

    /**
     * return the Title associated with this Axis (through the linked TextDisp record)
     *
     * @return
     */
    val label: String
        @Deprecated("use getTitle")
        get() = title

    /**
     * set the Title associated with this Axis (through the linked TextDisp record)
     *
     * @param l
     */
    // finds or adds TextDisp records for this axis (for labels and fonts)
    var title: String?
        get() {
            if (linkedtd == null)
                getTitleTD(false)
            return if (linkedtd != null) linkedtd!!.toString() else ""
        }
        set(l) {
            if (l == null) {
                this.removeTitle()
                return
            }
            if (linkedtd == null) {
                getTitleTD(true)
            }
            linkedtd!!.setText(l)
        }


    /**
     * return the Font object associated with the Axis Title, or null if  none
     */
    val titleFont: Font?
        get() {
            try {
                if (linkedtd != null) {
                    val fx = Chart.findRec(linkedtd!!.chartArr, Fontx::class.java) as Fontx
                    return this.parentChart!!.workBook!!.getFont(fx.ifnt)
                }
            } catch (e: Exception) {
            }

            return null
        }

    /**
     * returns the font of the Axis Title
     */
    override// finds or adds TextDisp records for this axis (for labels and fonts)
    val font: io.starter.formats.XLS.Font?
        get() {
            if (linkedtd == null) {
                getTitleTD(true)
            }
            val idx = linkedtd!!.fontId
            return this.workBook!!.getFont(idx)
        }

    /**
     * return the Font object used for Axis labels
     *
     * @return
     */
    val labelFont: io.starter.formats.XLS.Font?
        get() {
            try {
                val fx = Chart.findRec(chartArr, Fontx::class.java) as Fontx
                return this.parentChart!!.workBook!!.getFont(fx.ifnt)
            } catch (e: NullPointerException) {
                return this.parentChart!!.defaultFont
            }

        }

    /**
     * returns true if this axis is reversed
     * <br></br>If horizontal axis, default= on bottom, reversed= on top
     * <br></br>If vertical axis, default= LHS, reversed= RHS
     *
     * @return
     */
    // shouldn't
    val isReversed: Boolean
        get() {
            if (axis.toInt() == ChartConstants.XAXIS) {
                val c = getCatserRange(false)
                if (c != null)
                    return c.isReversed
            }
            val v = Chart.findRec(this.chartArr, ValueRange::class.java) as ValueRange
            return v?.isReversed ?: false
        }

    // see if have series-specific formats
    // see if it's a value X axis with a custom
    val numberFormat: String?
        get() {
            val f = Chart.findRec(this.chartArr, Ifmt::class.java) as Ifmt
            var i = 0
            if (f != null) {
                i = f.fmt
            } else {
                val s = this.parentChart!!.getAllSeries(-1)
                if (s.size > 0) {
                    return if (axis.toInt() == ChartConstants.YAXIS) {
                        (s[0] as Series).seriesFormatPattern
                    } else
                        (s[0] as Series).categoryFormatPattern
                }
            }
            return "General"
        }

    protected// shouldn't
    val minMax: DoubleArray
        get() {
            val v = Chart.findRec(this.chartArr, ValueRange::class.java) as ValueRange
            return if (v != null) {
                doubleArrayOf(v.min, v.max)
            } else doubleArrayOf(0.0, 0.0)

        }

    /**
     * return the major tick unit of this Y or Value axis
     *
     * @return
     */
    // try a default
    val majorUnit: Int
        get() {
            val v = Chart.findRec(this.chartArr, ValueRange::class.java) as ValueRange
            return v?.majorTick?.toInt() ?: 10
        }

    /**
     * return the minor tick unit of this Y or Value axis
     *
     * @return
     */
    // try a default
    val minorUnit: Int
        get() {
            val v = Chart.findRec(this.chartArr, ValueRange::class.java) as ValueRange
            return v?.minorTick?.toInt() ?: 0
        }

    /**
     * returns true if either Automatic min or max scale is set for the Y or Value axis
     *
     * @return
     */
    /**
     * sets the automatic scale option on or off for the Y or Value axis
     * <br></br>Automatic Scaling will automatically set the scale maximum, minimum and tick units
     *
     * @param b
     */
    var isAutomaticScale: Boolean
        get() {
            val v = Chart.findRec(this.chartArr, ValueRange::class.java) as ValueRange
            return if (v != null) {
                v.isAutomaticMin || v.isAutomaticMax
            } else false
        }
        set(b) {
            val v = Chart.findRec(this.chartArr, ValueRange::class.java) as ValueRange
            if (v != null) {
                v.isAutomaticMin = b
                v.isAutomaticMax = b
            }
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
     * @return int - one of the Axis Label placement constants above
     */
    /**
     * sets the axis labels position or placement to the desired value (these match Excel placement options)
     *
     * Possible options:
     * <br></br>Axis.INVISIBLE - hides the axis
     * <br></br>Axis.LOW - low end of plot area
     * <br></br>Axis.HIGH - high end of plot area
     * <br></br>Axis.NEXTTO- next to axis (default)
     *
     * @param Placement - int one of the Axis placement constants listed above
     */
    // shoudn't
    // shoudn't
    var axisPlacement: Int
        get() {
            val t = Chart.findRec(this.chartArr, Tick::class.java) as Tick
            if (t != null) {
                val p = t.getOption("tickLblPos")
                if (p == null || p == "none")
                    return Axis.INVISIBLE
                else if (p == "low")
                    return Axis.LOW
                else if (p == "high")
                    return Axis.HIGH
                else if (p == "nextTo")
                    return Axis.NEXTTO
            }
            return Axis.INVISIBLE
        }
        set(Placement) {
            val t = Chart.findRec(this.chartArr, Tick::class.java) as Tick
            if (t != null) {
                when (Placement) {
                    Axis.INVISIBLE -> t.setOption("tickLblPos", "none")
                    Axis.LOW -> t.setOption("tickLblPos", "low")
                    Axis.HIGH -> t.setOption("tickLblPos", "high")
                    Axis.NEXTTO -> t.setOption("tickLblPos", "nextTo")
                }
            }
        }

    override fun init() {
        super.init()
        axis = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
    }

    fun setAxis(wType: Int) {
        var wType = wType
        this.axis = wType.toShort()
        if (wType == ChartConstants.XVALAXIS) wType = ChartConstants.XAXIS        // 20090108 KSC: XVALAXIS is type of X axis with VAL records
        val b = ByteTools.shortToLEBytes(wType.toShort())
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    /**
     * remove the title for this axis
     */
    fun removeTitle() {
        if (linkedtd != null) {
            var x = Chart.findRecPosition(ap!!.chartArr, TextDisp::class.java)
            while (x > -1) {
                if ((ap!!.chartArr[x] as TextDisp).type == TextDisp.convertType(axis.toInt())) {
                    ap!!.chartArr.removeAt(x)
                    break
                }
                x++
                if ((ap!!.chartArr[x] as BiffRec).opcode != XLSConstants.TEXTDISP)
                    break
            }
        }
        linkedtd = null
    }

    /**
     * returns true if plot area has a bonding box/border
     *
     * @return
     */
    fun hasPlotAreaBorder(): Boolean {
        val f = Chart.findRec(ap!!.chartArr, Frame::class.java) as Frame
        return f?.hasBox() ?: false
    }

    /**
     * link this axis with it's parent
     *
     * @param a
     */
    fun setAP(a: AxisParent) {
        ap = a
    }

    /**
     * finds or adds TextDisp records for this axis (for labels and fonts)
     */
    private fun getTitleTD(add: Boolean) {
        val tdtype = TextDisp.convertType(axis.toInt())
        var pos = 0
        // Must add if can't find existing label
        var x = Chart.findRecPosition(ap!!.chartArr, TextDisp::class.java)
        pos = x
        while (x > 0 && linkedtd == null) {
            val td = ap!!.chartArr[x] as TextDisp
            if (td.type == tdtype)
            // found it
                linkedtd = td
            else {
                if (ap!!.chartArr[++x] !is TextDisp)
                    x = -1
            }
        }
        if (linkedtd == null && add) {
            linkedtd = TextDisp.getPrototype(tdtype, "", this.wkbook) as TextDisp
            if (pos < 0)
                pos = Chart.findRecPosition(ap!!.chartArr, PlotArea::class.java)
            else {    // otherwise, TextDisp(s) exist already; position correctly
                // set font to other axis font as default
                val td = ap!!.chartArr[pos] as TextDisp
                linkedtd!!.fontId = td.fontId
                if (axis.toInt() != ChartConstants.XAXIS)
                    pos++
            }
            if (pos < 0)
                pos = Chart.findRecPosition(ap!!.chartArr, ChartFormat::class.java)
            linkedtd!!.parentChart = this.parentChart
            ap!!.chartArr.add(pos, linkedtd)
        }
    }

    /**
     * set the font index for this Axis (for title)
     *
     * @param fondId
     */
    fun setFont(fondId: Int) {
        if (linkedtd == null) {
            getTitleTD(true)    // finds or adds TextDisp records for this axis (for labels and fonts)
        }
        linkedtd!!.fontId = fondId
    }

    /**
     * utility method to find the CatserRange rec associated with this Axis
     *
     * @return
     */
    protected fun getCatserRange(bCreate: Boolean): CatserRange {
        var csr = Chart.findRec(chartArr, CatserRange::class.java) as CatserRange
        if (csr == null) {
            csr = CatserRange.prototype as CatserRange?
            csr.parentChart = this.parentChart
            chartArr.add(0, csr)
        }
        return csr
    }

    /**
     * returns true of this axis displays major gridlines
     *
     * @return
     */
    protected fun hasGridlines(type: Int): Boolean {
        var j = Chart.findRecPosition(chartArr, AxisLineFormat::class.java)
        if (j != -1) {
            try {
                while (j < chartArr.size) {
                    val al = chartArr[j] as AxisLineFormat
                    val id = al.id.toInt()
                    if (id == type)
                    // Y Axis draw gridline usually true, trap if do not have (see below)
                        return true
                    j += 2    // Skip line format
                }
            } catch (e: ClassCastException) {
            }

        }
        return false
    }

    /**
     * returns the SVG necessary to describe the desired line (referenced by id)
     *
     * @param id @see AxisLineFormat id types
     * @return
     */
    protected fun getLineSVG(id: Int): String {
        val lf = getAxisLine(id)
        return if (lf != null) lf.svg else ""
    }

    /**
     * return the lineformat rec for the given axis line type
     *
     * @param type AxisLineFormat type
     * @return Line Format
     * @see AxisLineFormat.ID_MAJOR_GRID, et. atl.
     */
    protected fun getAxisLine(type: Int): LineFormat? {
        val j = getAxisLineFormat(type)
        return if (j > -1) chartArr[j + 1] as LineFormat else null

    }

    /**
     * return the AxisLineFormat of the desired type, or create if none and bCreate
     *
     * @param type    AxisLineFormat Type:  AxisLineFormat.ID_AXIS_LINE, AxisLineFormat.ID_MAJOR_GRID, AxisLineFormat.ID_MINOR_GRID, AxisLineFormat.ID_WALLORFLOOR
     * @param bCreate
     * @return
     */
    protected fun getAxisLineFormat(type: Int, bCreate: Boolean): AxisLineFormat? {
        var alf: AxisLineFormat? = null
        var j = Chart.findRecPosition(chartArr, AxisLineFormat::class.java)
        if (j == -1 && !bCreate)
            return null
        if (j > -1) {
            try {
                while (j < chartArr.size) {
                    alf = chartArr[j] as AxisLineFormat
                    if (alf.id.toInt() == type)
                        return alf
                    else if (alf.id > type)
                        break
                    j += 2
                }
            } catch (e: ClassCastException) {
            }

        }
        j = 1
        while (j < chartArr.size) {
            val b = chartArr[j]
            if (b.opcode == XLSConstants.AREAFORMAT ||
                    b.opcode == XLSConstants.GELFRAME ||
                    b.opcode.toInt() == 2213 ||    /* TextPropsStream */
                    b.opcode.toInt() == 2212 ||  /* ShapePropsStream */
                    b.opcode.toInt() == 2206)
            /* CtrlMlFrt */
                break
            j++
        }
        alf = AxisLineFormat.prototype as AxisLineFormat?
        alf!!.setId(type)    // default has major gridlines
        chartArr.add(j++, alf)
        alf.parentChart = this.parentChart
        val lf = LineFormat.prototype as LineFormat?
        lf!!.parentChart = this.parentChart
        chartArr.add(j, lf)
        return alf

    }

    /**
     * return the index of the AxisLineFormat of the desired type, -1 if none
     *
     * @param type AxisLineFormat Type:  AxisLineFormat.ID_AXIS_LINE, AxisLineFormat.ID_MAJOR_GRID, AxisLineFormat.ID_MINOR_GRID, AxisLineFormat.ID_WALLORFLOOR
     * @return
     */
    protected fun getAxisLineFormat(type: Int): Int {
        var alf: AxisLineFormat? = null
        var j = Chart.findRecPosition(chartArr, AxisLineFormat::class.java)
        if (j == -1) return j
        try {
            while (j < chartArr.size) {
                alf = chartArr[j] as AxisLineFormat
                if (alf.id.toInt() == type)
                    return j
                else if (alf.id > type)
                    break
                j += 2
            }
        } catch (e: ClassCastException) {
        }

        return if (j > chartArr.size) -1 else j    // not found

    }

    /**
     * gets or creates YMult (value multiplier) record
     *
     * @param bCreate
     * @return
     */
    protected fun getYMultRec(bCreate: Boolean): YMult? {
        var ym = Chart.findRec(chartArr, YMult::class.java) as YMult
        if (ym == null && bCreate) {
            ym = YMult.prototype as YMult?
            ym.parentChart = this.parentChart
            chartArr.add(1, ym)    // 2nd, after ValueRange
        }
        return ym
    }

    /**
     * return the JSON/Dojo representation of this axis
     * chartType int necessary for parsing AXIS options: horizontal charts "switch" axes ...
     * All options have NOT been gathered at this point
     *
     * @return JSONObject
     */
    fun getJSON(wbh: io.starter.OpenXLS.WorkBookHandle, chartType: Int, yMax: Double, yMin: Double, nSeries: Int): JSONObject {
        val axisJSON = JSONObject()
        val axisOptions = JSONObject()
        try {
            if (axis.toInt() == ChartConstants.YAXIS && chartType != ChartConstants.BARCHART) {
                axisOptions.put("vertical", true)
            } else if (axis.toInt() == ChartConstants.XAXIS && chartType == ChartConstants.BARCHART)
                axisOptions.put("vertical", true)

            // 20090721 KSC: dojo 1.3.1 has label element
            axisOptions.put("label", title)

            // TODO: Dojo Axis Options:
            // fixLower, fixUpper	("minor"/"major")
            // includeZero, natural (true)
            // min, max
            // minorTicks, majorTicks, microTicks (true/false)
            // minorTickStep, majorTickStep, microTickStep (#)
            for (i in chartArr.indices) {
                val b = chartArr[i]
                if (b is CatserRange) {
// for x/Category axis:  if has labels, gather and input into axis label JSON array
                    val categories = this.parentChart!!.getCategories(-1)
                    if (categories != null) {
                        // Category Labels
                        val labels = JSONArray()
                        if (b.crossBetween) { // categories appear mid-axis so put "spacers" at 0 and max
                            axisOptions.put("includeZero", true)        // always true????
                            val nullCat = JSONObject()
                            nullCat.put("value", 0)
                            nullCat.put("text", "")
                            labels.put(nullCat)
                        }
                        val cats = CellRange.getValuesAsJSON(categories[0], wbh)    // parse category range into JSON Array
                        for (z in 0 until cats.length()) {
                            val aCat = JSONObject()
                            aCat.put("value", z + b.catLabel)    // should= +1 i.e. a category label appears with each category
                            aCat.put("text", cats.get(z))
                            labels.put(aCat)
                        }
                        if (b.crossBetween) { // categories appear mid-axis so put "spacers" at 0 and max
                            val nullCat = JSONObject()
                            nullCat.put("value", cats.length() + b.catLabel)
                            nullCat.put("text", "")
                            labels.put(nullCat)
                            axisOptions.put("max", cats.length() + b.catLabel/*+1*/)
                        }
                        axisOptions.put("labels", labels)
                        // Defaults axis values ...
                        axisOptions.put("fixLower", "major")        // what do these do??
                        axisOptions.put("fixUpper", "major")
                    }

                    //	            	if (c.getCatCross()!=1)
                    if (b.catMark != 1) {
                        /*	            		The catMark field defines how often tick marks appear along the category or series axis. A value of 01 indicates
	            		that a tick mark will appear between each category or series; a value of 02 means a label appears between every
	            		other category or series, etc.
*/
                    }

                    //	            	if (c.getCrossMax())
                } else if (b is ValueRange) {
                    if (axis.toInt() == ChartConstants.YAXIS)
                    // normal
                        b.setMaxMin(yMax, yMin) // must do first
                    else
                        b.setMaxMin(nSeries.toDouble(), 0.0)    // scatter and bubble charts have X axis with Value Range, scale is 0 to number of series

                    // y major/minor scales
                    axisOptions.put("min", b.min)
                    axisOptions.put("max", b.max)
                    axisOptions.put("majorTickStep", b.majorTick)
                } else if (b is AxisLineFormat) {
                    val gridJSON = JSONObject()
                    gridJSON.put("type", "Grid")
                    val id = b.id.toInt()
                    when (id) {
                        AxisLineFormat.ID_MAJOR_GRID -> if (axis.toInt() == ChartConstants.XAXIS || chartType == ChartConstants.BARCHART)
                            gridJSON.put("hMajorLines", false)
                        else
                            gridJSON.put("vMajorLines", false)
                        AxisLineFormat.ID_MINOR_GRID -> {
                        }
                    }
                    axisJSON.put("back_grid", gridJSON)
                }
            }
            if (axis.toInt() == ChartConstants.YAXIS)
                axisJSON.put("y", axisOptions)
            if (axis.toInt() == ChartConstants.XAXIS)
                axisJSON.put("x", axisOptions)
        } catch (e: JSONException) {
            Logger.logErr("Error getting Axis JSON: $e")
        }

        return axisJSON
    }

    /**
     * interface for setting Axis rec-specific XML options
     * in a generic fashion
     *
     * @see OpenXLS.handleChartElement
     *
     * @see ChartHandle.getXML
     */
    override fun setChartOption(op: String, `val`: String): Boolean {
        if (op.equals("Label", ignoreCase = true)) {
            this.title = `val`
            return true
        } else if (op.equals("CatCross", ignoreCase = true)) {
            getCatserRange(true).catCross = Integer.parseInt(`val`)
            return true
        } else if (op.equals("LabelCross", ignoreCase = true)) {
            getCatserRange(true).catLabel = Integer.parseInt(`val`)
            return true
        } else if (op.equals("Marks", ignoreCase = true)) {
            getCatserRange(true).catMark = Integer.parseInt(`val`)
            return true
        } else if (op.equals("CrossBetween", ignoreCase = true)) {
            getCatserRange(true).crossBetween = `val` == "true"
            return true
        } else if (op.equals("CrossMax", ignoreCase = true)) {
            getCatserRange(true).crossMax = `val` == "true"
            return true
        } else if (op.equals("MajorGridLines", ignoreCase = true)) {
            if (`val` == "false") {
                val j = getAxisLineFormat(AxisLineFormat.ID_MAJOR_GRID)
                if (j > -1) {
                    chartArr.removeAt(j)    // remove AxisLineFormat
                    chartArr.removeAt(j)    // remove corresponding Line Format
                }
            } else {    // add major grid lines
                getAxisLineFormat(AxisLineFormat.ID_MAJOR_GRID, true)
            }
        } else if (op.equals("MinorGridLines", ignoreCase = true)) {
            if (`val` == "false") {
                val j = getAxisLineFormat(AxisLineFormat.ID_MINOR_GRID)
                if (j > -1) {
                    chartArr.removeAt(j)    // remove AxisLineFormat
                    chartArr.removeAt(j)    // remove corresponding Line Format
                }
            } else {    // add major grid lines
                getAxisLineFormat(AxisLineFormat.ID_MINOR_GRID, true)
            }
        } else if (op.equals("AddArea", ignoreCase = true)) {
            if (axis.toInt() == ChartConstants.XAXIS) {
                val alf0 = AxisLineFormat.prototype as AxisLineFormat?
                alf0!!.setId(AxisLineFormat.ID_AXIS_LINE)
                this.addChartRecord(alf0)
                val lf0 = LineFormat.getPrototype(0, 0) as LineFormat
                this.addChartRecord(lf0)
            }
            val alf = AxisLineFormat.prototype as AxisLineFormat?
            alf!!.setId(AxisLineFormat.ID_WALLORFLOOR)
            this.addChartRecord(alf)
            val lf = LineFormat.getPrototype(0, -1) as LineFormat
            if (axis.toInt() == 1)
                lf.lineStyle = 5    //none
            this.addChartRecord(lf)
            val af = AreaFormat.getPrototype(axis.toInt()) as AreaFormat
            this.addChartRecord(af)
            return true
        } else if (op == "AreaFg") {    // custom foreground on Wall, Side or Floor
            val af = Chart.findRec(chartArr, AreaFormat::class.java) as AreaFormat
            af.seticvFore(Integer.valueOf(`val`).toInt())
            return true
        } else if (op == "AreaBg") {    // custom bg on Wall, SIde or Floor
            val af = Chart.findRec(chartArr, AreaFormat::class.java) as AreaFormat
            af.seticvBack(Integer.valueOf(`val`).toInt())
            return true
        } else if (linkedtd != null) { // see if associated TextDisp can handle
            return linkedtd!!.setChartOption(op, `val`)
        } else if (op.equals("MajorGridLines", ignoreCase = true)) {
            if (`val` == "false") {
                val j = getAxisLineFormat(AxisLineFormat.ID_MAJOR_GRID)
                if (j > -1) {
                    chartArr.removeAt(j)    // remove AxisLineFormat
                    chartArr.removeAt(j)    // remove corresponding Line Format
                }
            } else {    // add major grid lines
                getAxisLineFormat(AxisLineFormat.ID_MAJOR_GRID, true)
            }
        } else if (op.equals("MinorGridLines", ignoreCase = true)) {
            if (`val` == "false") {
                val j = getAxisLineFormat(AxisLineFormat.ID_MINOR_GRID)
                if (j > -1) {
                    chartArr.removeAt(j)    // remove AxisLineFormat
                    chartArr.removeAt(j)    // remove corresponding Line Format
                }
            } else {    // add major grid lines
                getAxisLineFormat(AxisLineFormat.ID_MINOR_GRID, true)
            }
        }
        return false
    }

    /**
     * sets a specific OOXML axis option; most options apply to any type of axis (Cat, Value, Ser, Date)
     * <br></br>can be one of:
     * <br></br>axPos		-   position of the axis (b, t, r, l)
     * <br></br>crosses			possible crossing points (autoZero, max, min)
     * <br></br>crossBeteween	whether axis crosses the cat. axis between or on categories (value axis only)  (between, midCat)
     * <br></br>crossesAt		where on axis the perpendicular axis crosses (double val)
     * <br></br>lblAlign			text alignment for tick labels (ctr, l, r) (cat only)
     * <br></br>lblOffset		distance of labels from the axis (0-1000)  (cat only)
     * <br></br>majorTickMark	major tick mark position (cross, in, none, out)
     * <br></br>minorTickMark	minor tick mark position ("")
     * <br></br>tickLblPos		tick label position (high, low, nextTo, none)
     * <br></br>tickLblSkip		how many tick labels to skip between label (int >= 1)	(cat only)
     * <br></br>tickMarkSkip		how many tick marks to skip betwen ticks (int >= 1)		(cat only)
     * <br></br>majorUnit		distance between major tick marks (val, date ax only) (double >= 0)
     * <br></br>minorUnit		distance between minor tick marks (val, date ax only) (double >= 0)
     * <br></br>MajorGridLines
     * <br></br>MinorGridLines
     *
     * @param op
     * @param val
     */
    private fun setOption(op: String, `val`: String) {
        if (op == "axPos") {                    // val= "b" (bottom) "l", "t", "r"   -->?????
            axPos = `val`    // for now
        } else if (op == "lblOffset" || op == "lblAlgn") {
            var cl = Chart.findRec(chartArr, CatLab::class.java) as CatLab
            if (cl == null) {
                cl = CatLab.prototype as CatLab?
                cl.parentChart = this.parentChart
                chartArr.add(1, cl)    // second in chart array, after CatSerRange
            }
            cl.setOption(op, `val`)
        } else if (op == "tickLblPos" ||        // val= high (=at high end of perp. axis), low (=at low end of perp. axis), nextTo, none (=no axis labels) Tick

                op == "majorTickMark" ||    // major tick marks (cross, in, none, out)

                op == "minorTickMark") {    // minor tick marks (cross, in, none, out)
            val t = Chart.findRec(chartArr, Tick::class.java) as Tick
            t.setOption(op, `val`)
        } else if (op.equals("MajorGridLines", ignoreCase = true)) {
            if (`val` == "false") {
                val j = getAxisLineFormat(AxisLineFormat.ID_MAJOR_GRID)
                if (j > -1) {
                    chartArr.removeAt(j)    // remove AxisLineFormat
                    chartArr.removeAt(j)    // remove corresponding Line Format
                }
            } else {    // add major grid lines
                getAxisLineFormat(AxisLineFormat.ID_MAJOR_GRID, true)
            }
        } else if (op.equals("MinorGridLines", ignoreCase = true)) {
            if (`val` == "false") {
                val j = getAxisLineFormat(AxisLineFormat.ID_MINOR_GRID)
                if (j > -1) {
                    chartArr.removeAt(j)    // remove AxisLineFormat
                    chartArr.removeAt(j)    // remove corresponding Line Format
                }
            } else {    // add major grid lines
                getAxisLineFormat(AxisLineFormat.ID_MINOR_GRID, true)
            }
        } else {        // valuerange, caterrange options	-- crosses, crossBetween, crossesAt, tickMarkSkip (cat only), tickLblSkip (cat only), majorUnit (val only), minorUnit (val only)
            // KSC: TESTING
            //Logger.logInfo("Setting option: " + op + "=" + val);
            for (i in chartArr.indices) {
                val b = chartArr[i]
                if (b is CatserRange) {
                    if (b.setOption(op, `val`))
                        break
                } else if (b is ValueRange) {
                    if (b.setOption(op, `val`))
                        break
                }
            }
        }
    }

    /**
     * get the desired axis option for this axis
     * <br></br>can be one of:
     * <br></br>axPos		-   position of the axis (b, t, r, l)
     * <br></br>crosses			possible crossing points (autoZero, max, min)
     * <br></br>crossBeteween	whether axis crosses the cat. axis between or on categories (value axis only)  (between, midCat)
     * <br></br>crossesAt		where on axis the perpendicular axis crosses (double val)
     * <br></br>lblAlign			text alignment for tick labels (ctr, l, r) (cat only)
     * <br></br>lblOffset		distance of labels from the axis (0-1000)  (cat only)
     * <br></br>majorTickMark	major tick mark position (cross, in, none, out)
     * <br></br>minorTickMark	minor tick mark position ("")
     * <br></br>tickLblPos		tick label position (high, low, nextTo, none)
     * <br></br>tickLblSkip		how many tick labels to skip between label (int >= 1)	(cat only)
     * <br></br>tickMarkSkip		how many tick marks to skip betwen ticks (int >= 1)		(cat only)
     * <br></br>majorUnit		distance between major tick marks (val, date ax only) (double >= 0)
     * <br></br>minorUnit		distance between minor tick marks (val, date ax only) (double >= 0)
     *
     * @param op String option name
     * @return String val of option or ""
     */
    fun getOption(op: String): String? {
        if (op == "axPos") {                    // val= "b" (bottom) "l", "t", "r"   -->?????
            return axPos    // for now -- can't find matching Axis attribute
        } else if (op == "lblAlign" || op == "lblOffset") {
            val c = Chart.findRec(chartArr, CatLab::class.java) as CatLab
            return c?.getOption(op)
// use defaults
        } else if (op == "crossesAt" ||        // specifies where axis crosses		  -- numCross or catCross

                op == "orientation" ||        // axis orientation minMax or maxMin  -- fReverse

                op == "crosses" ||            // specifies how axis crosses it's perpendicular axis (val= max, min, autoZero)  -- fbetween + fMaxCross?/fAutoCross + fMaxCross

                op == "max" ||                // axis max - valueRange only?

                op == "max" ||                // axis min- valueRange only?

                op == "tickLblSkip" ||    //val= how many tick labels to skip btwn label -- catLabel -- Catserrange only??

                op == "tickMarkSkip" ||    //val= how many tick marks to skip before next one is drawn -- catMark -- catsterrange only?

                op == "crossBetween") {    // value axis only -- val= between, midCat, crossBetween
            // logScale-- ValueRange
            for (i in chartArr.indices) {
                val b = chartArr[i]
                if (b is CatserRange) {
                    return b.getOption(op)
                } else if (b is ValueRange) {
                    return b.getOption(op)
                }
            }
            //TICK Options
        } else if (op == "tickLblPos" ||        // val= high (=at high end of perp. axis), low (=at low end of perp. axis), nextTo, none (=no axis labels) Tick

                op == "majorTickMark" ||    // major tick marks (cross, in, none, out)

                op == "minorTickMark") {    // minor tick marks (cross, in, none, out)
            val t = Chart.findRec(chartArr, Tick::class.java) as Tick
            return t.getOption(op)
        }
        return null
    }

    /**
     * return the JSON/Dojo representation of this axis
     * chartType int necessary for parsing AXIS options: horizontal charts "switch" axes ...
     *
     * @return JSONObject
     */
    fun getMinMaxJSON(wbh: io.starter.OpenXLS.WorkBookHandle, chartType: Int, yMax: Double, yMin: Double, nSeries: Int): JSONObject {
        val axisJSON = JSONObject()
        val axisOptions = JSONObject()
        try {
            for (i in chartArr.indices) {
                val b = chartArr[i]
                if (b is CatserRange) {
// for x/Category axis:  if has labels, gather and input into axis label JSON array
                    val categories = this.parentChart!!.getCategories(-1)
                    if (categories != null) {
                        val cats = CellRange.getValuesAsJSON(categories[0], wbh)    // parse category range into JSON Array
                        axisOptions.put("max", cats.length() + b.catLabel/*+1*/)
                    }
                    break    // only need this per axis
                } else if (b is ValueRange) {
                    if (axis.toInt() == ChartConstants.YAXIS)
                    // normal
                        b.setMaxMin(yMax, yMin)    // must do first
                    else
                        b.setMaxMin(nSeries.toDouble(), 0.0)    // scatter and bubble charts have X axis with Value Range, scale is 0 to number of series

                    // y major/minor scales
                    axisOptions.put("min", b.min)
                    axisOptions.put("max", b.max)
                    axisOptions.put("majorTickStep", b.majorTick)
                    break    // only need this per axis
                }
            }
            if (axis.toInt() == ChartConstants.YAXIS)
                axisJSON.put("y", axisOptions)
            if (axis.toInt() == ChartConstants.XAXIS)
                axisJSON.put("x", axisOptions)
        } catch (e: JSONException) {
            Logger.logErr("Error getting Axis JSON: $e")
        }

        return axisJSON
    }

    /**
     * return the OOXML txPr element for this axis
     *
     * @return
     */
    fun gettxPr(): TxPr? {
        return txpr
    }

    /**
     * set the OOXML title element for this axis
     *
     * @param t
     */
    fun settxPr(t: TxPr) {
        txpr = t.cloneElement() as TxPr
    }

    override fun toString(): String {
        var s = ""
        when (axis) {
            ChartConstants.XAXIS -> s = "XAxis"
            ChartConstants.YAXIS -> s = "YAxis"
            ChartConstants.ZAXIS -> s = "ZAxis"
            ChartConstants.XVALAXIS -> s = "XValAxis"
        }
        if (linkedtd != null)
            s = s + " " + linkedtd!!.toString()
        return s
    }

    /**
     * return the maximum value of this Value or Y axis scale
     *
     * @return
     */
    fun getMaxScale(minmax: DoubleArray): Double {
        val v = Chart.findRec(this.chartArr, ValueRange::class.java) as ValueRange
        if (v != null) { // shouldn't
            if (v.isAutomaticMax)
                v.setMaxMin(minmax[1], minmax[0])
            return v.max
        }
        return -1.0
    }

    /**
     * return the minimum value of this value or Y axis scale
     *
     * @return
     */
    fun getMinScale(minmax: DoubleArray): Double {
        /*
         * Because a horizontal (category) axis (axis: A line bordering the chart plot area used as
         * a frame of reference for measurement.  * The y axis is usually the vertical axis and contains data. The x-axis is usually the horizontal axis and contains categories.)
         * displays text labels instead of numeric intervals,
         * there are fewer scaling options that you can change than there are for a vertical (value) axis.
         * However, you can change the number of categories to display between tick marks, the order in which
         * to display categories, and the point where the two axes cross.
         */
        val v = Chart.findRec(this.chartArr, ValueRange::class.java) as ValueRange
        if (v != null) {
            /*
			 if (wType==YAXIS)	// normal
            	v.setMaxMin(yMax, yMin); 	// must do first
            else
            	v.setMaxMin(nSeries, 0);	// scatter and bubble charts have X axis with Value Range, scale is 0 to number of series
			 */
            // why? should already			v.setParentChart(this.getParentChart());
            if (v.isAutomaticMin)
                v.setMaxMin(minmax[1], minmax[0])
            return v.min
        }
        return -1.0
    }

    /**
     * set the minimum value of this axis scale
     * <br></br>Note: this disables automatic scaling
     *
     * @param Min
     */
    fun setMinScale(min: Int) {
        // TODO: also update ticks?
        val v = Chart.findRec(this.chartArr, ValueRange::class.java) as ValueRange
        if (v != null) { // shouldn't
            v.min = min
        }
    }

    /**
     * set the maximum value of this axis scale
     * <br></br>Note: this disables automatic scaling
     *
     * @param Max
     */
    fun setMaxScale(max: Int) {
        // TODO: also update ticks?
        val v = Chart.findRec(this.chartArr, ValueRange::class.java) as ValueRange
        if (v != null) { // shouldn't
            v.max = max
        }
    }

    /**
     * parse OOXML axis element
     *
     * @param xpp     XmlPullParser positioned at correct elemnt
     * @param axisTag catAx, valAx, serAx, dateAx
     * @param lastTag Stack of element names
     */
    // noMultiLvlLbl -- val= 1 (true) means draw labels as flat text; not included or 0 (false)= draw labels as a heirarchy
    fun parseOOXML(xpp: XmlPullParser, axisTag: String, lastTag: Stack<String>, bk: WorkBookHandle) {
        // crossAx -- need to parse?
        // auto -- need to parse?
        try {
            var eventType = xpp.eventType
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    val tnm = xpp.name
                    if (tnm == "scaling") {        // additional axis settings
                        lastTag.push(tnm)
                        val sc = Scaling.parseOOXML(xpp, lastTag) as Scaling
                        var s = sc.getOption("orientation")
                        if (s != null) this.setOption("orientation", s)
                        s = sc.getOption("min")
                        if (s != null) this.setOption("min", s)
                        s = sc.getOption("max")
                        if (s != null) this.setOption("max", s)
                        // the below children only have 1 attribute: val
                    } else if (tnm == "axPos") {        // // position of the axis (b, t, r, l)
                        this.setOption(tnm, xpp.getAttributeValue(0))
                    } else if (tnm == "majorGridlines" || tnm == "minorGridlines") {
                        lastTag.push(tnm)
                        parseGridlinesOOXML(xpp, lastTag, bk)
                    } else if (tnm == "title") {
                        lastTag.push(tnm)
                        this.ooxmlTitle = Title.parseOOXML(xpp, lastTag, bk).cloneElement() as Title
                        this.title = this.ooxmlTitle!!.title
                    } else if (tnm == "numFmt") {
                        this.nf = NumFmt.parseOOXML(xpp).cloneElement() as NumFmt
                    } else if (tnm == "majorTickMark" ||    // major tick mark position (cross, in, none, out)

                            tnm == "minorTickMark" ||        // minor tick mark position ("")

                            tnm == "tickLblPos") {        // tick label position (high, low, nextTo, none)
                        this.setOption(tnm, xpp.getAttributeValue(0))
                    } else if (tnm == "spPr") {    // axis shape properties - for axis or gridlines
                        lastTag.push(tnm)
                        this.spPr = SpPr.parseOOXML(xpp, lastTag, bk).cloneElement() as SpPr
                    } else if (tnm == "txPr") {        // text Properties for axis
                        lastTag.push(tnm)
                        this.settxPr(TxPr.parseOOXML(xpp, lastTag, bk).cloneElement() as TxPr)
                        // crossesAx = crossing axis id - need ?
                    } else if (tnm == "crosses" ||            // possible crossing points (autoZero, max, min)
                            tnm == "crossesAt") {            // where on axis the perpendicular axis crosses (double val)
                        this.setOption(tnm, xpp.getAttributeValue(0))
                    } else if (tnm == "crossBetween") {        // whether axis crosses the cat. axis between or on categories (value axis only)  (between, midCat)
                        this.setOption(tnm, xpp.getAttributeValue(0))
                    } else if (// cat, date, ser ax only
                    // auto -- Date only?
                            tnm == "lblAlign" ||            // text alignment for tick labels (ctr, l, r)  only for cat

                            tnm == "lblOffset" ||            // distance of labels from the axis (0-1000)	only for cat, date

                            tnm == "tickLblSkip" ||        // how many tick labels to skip between label (int >= 1)

                            tnm == "tickMarkSkip") {        // how many tick marks to skip betwen ticks (int >= 1)
                        this.setOption(tnm, xpp.getAttributeValue(0))
                        // TODO: noMultiLvlLbl
                    } else if (// val, ser ax + some date ax
                            tnm == "crossBeteween" ||        // whether axis crosses the cat. axis between or on categories (value axis only)  (between, midCat)

                            tnm == "majorUnit" ||            // distance between major tick marks (val, date ax only) (double >= 0)

                            tnm == "minorUnit") {            // distance between minor tick marks (val, date ax only) (double >= 0)
                        this.setOption(tnm, xpp.getAttributeValue(0))
                    } else if (tnm == "dispUnits") {    // valAx only
                        parseDispUnitsOOXML(xpp, lastTag)
                    }
                    // TODO: date ax specifics
                } else if (eventType == XmlPullParser.END_TAG) {
                    val endTag = xpp.name
                    if (endTag == axisTag) {
                        lastTag.pop()
                        break
                    }
                }
                eventType = xpp.next()
            }
        } catch (e: Exception) {
            Logger.logErr("Axis: $e")
        }

    }

    /**
     * parse major or minor Gridlines element
     *
     * @param xpp
     * @param lastTag
     */
    private fun parseGridlinesOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle) {
        val endTag = lastTag.peek()
        this.setOption(endTag, "true")
        try {
            var eventType = xpp.eventType
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    val tnm = xpp.name
                    if (tnm == "spPr") {
                        lastTag.push(tnm)
                        val sppr = SpPr.parseOOXML(xpp, lastTag, bk).cloneElement() as SpPr
                        val lf = getAxisLine(if (endTag == "majorGridlines") AxisLineFormat.ID_MAJOR_GRID else AxisLineFormat.ID_MINOR_GRID)
                        lf!!.setFromOOXML(sppr)
                    }
                } else if (eventType == XmlPullParser.END_TAG) {
                    if (xpp.name == endTag) {
                        lastTag.pop()
                        break
                    }
                }
                eventType = xpp.next()
            }
        } catch (e: Exception) {
            Logger.logErr("parseGridLinesOOXML: $e")
        }

    }

    /**
     * parse the dispUnits child element of valAx
     * <br></br>TODO: do not know how to interpret most of these options
     *
     * @param xpp
     * @param lastTag
     */
    private fun parseDispUnitsOOXML(xpp: XmlPullParser, lastTag: Stack<String>) {
        try {
            var eventType = xpp.eventType
            var ym: YMult? = null
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    val tnm = xpp.name
                    if (tnm == "custUnit") {
                        ym = getYMultRec(true)
                        ym!!.customMultiplier = java.lang.Double.valueOf(xpp.getAttributeValue(0))
                    } else if (tnm == "builtInUnit") {
                        ym = getYMultRec(true)
                        ym!!.setAxMultiplierId(xpp.getAttributeValue(0))
                    }
                    // TODO: dispUnitLbl
                } else if (eventType == XmlPullParser.END_TAG) {
                    if (xpp.name == "dispUnits") {
                        lastTag.pop()
                        break
                    }
                }
                eventType = xpp.next()
            }
        } catch (e: Exception) {
            Logger.logErr("parseDispUnitsOOXML: $e")
        }

    }

    /**
     * generate the appropriate ooxml for the given axis
     *
     * @param type    0= Category Axis, 1= Value Axis, 2= Ser Axis, 3= Date Axis
     * @param id
     * @param crossId crossing axis id
     * @return ORDER:
     * Axis Common:
     * axisId		REQ
     * scaling	-- orientation
     * delete
     * axPos		REQ?
     * majorGridlines
     * minorGridlines
     * title
     * numFmt
     * majorTickMark
     * minorTickMark
     * tickLblPos
     * spPr
     * txPr
     * crossAx
     * crosses OR
     * crossesAt
     * after:
     * valAx:
     * crossBetween
     * majorUnit
     * minorUnit
     * dispUnits
     * catAx:
     * auto
     * lblAlign
     * lblOffet
     * tickLblSkip
     * tickMarkSkip
     * noMultiLvlLbl
     * serAx:
     * tickLblSkip
     * tickMarkSkip
     */

    fun getOOXML(type: Int, id: String, crossId: String): String {
        if (this.parentChart == null)
        // happens on ZAxis, XYValAxis ...
            this.parentChart = this.ap!!.parentChart
        val from2003 = !parentChart!!.workBook!!.isExcel2007

        val axisooxml = StringBuffer()
        var axis = ""
        when (type) {
            0    // cat axis "X"
            -> axis = "catAx"
            1    // val axis "Y"
            -> axis = "valAx"
            2 // xval axis "X" axis for multiple val axes: bubble & scatter charts
            -> axis = "valAx"
            3 // ser	 ("Z" axis)
            -> axis = "serAx"
            4  // date		TODO: Not correct or tested!
            -> axis = "dateAx"
        }
        // axis main element
        axisooxml.append("<c:$axis>")
        axisooxml.append("\r\n")
        // axId - required
        axisooxml.append("<c:axId val=\"$id\"/>")
        axisooxml.append("\r\n")
        // scaling - required
        var s = this.getOption("orientation")
        val d = this.minMax
        if (s != null || d[0] != d[1]) {    // if have orientation or min/max set ..
            axisooxml.append("<c:scaling>\r\n")
            if (s != null) axisooxml.append("<c:orientation val=\"$s\"/>\r\n")
            axisooxml.append("</c:scaling>\r\n")
        }
        // axPos - required
        if (this.getOption("axPos") != null) {
            axisooxml.append("<c:axPos val=\"" + this.getOption("axPos") + "\"/>")
            axisooxml.append("\r\n")
        } else {// it's required
            if (this.parentChart!!.chartType != ChartConstants.BARCHART) {
                if (axis == "catAx" || axis == "serAx")
                    axisooxml.append("<c:axPos val=\"b\"/>")
                else
                    axisooxml.append("<c:axPos val=\"l\"/>")
            } else {
                if (axis == "catAx" || axis == "serAx")
                    axisooxml.append("<c:axPos val=\"l\"/>")
                else
                    axisooxml.append("<c:axPos val=\"b\"/>")
            }
            axisooxml.append("\r\n")
        }
        // major Gridlines
        if (this.hasGridlines(AxisLineFormat.ID_MAJOR_GRID)) {
            axisooxml.append("<c:majorGridlines>")
            axisooxml.append(getAxisLine(AxisLineFormat.ID_MAJOR_GRID)!!.ooxml)
            axisooxml.append("</c:majorGridlines>\r\n")
        }
        // minor Gridlines
        if (this.hasGridlines(AxisLineFormat.ID_MINOR_GRID)) {
            axisooxml.append("<c:minorGridlines>")
            axisooxml.append(getAxisLine(AxisLineFormat.ID_MINOR_GRID)!!.ooxml)
            axisooxml.append("</c:minorGridlines>\r\n")
        }
        // Title
        if (this.ooxmlTitle != null)
            axisooxml.append(this.ooxmlTitle!!.ooxml)
        else if (from2003) {    // create OOXML title
            if (this.title != "") {
                val ttl = io.starter.formats.OOXML.Title(this.title)
                if (type == 0)
                    ttl.setLayout(.026, .378)
                else if (type == 1)
                    ttl.setLayout(.468, .863)
                axisooxml.append(ttl.ooxml)
            }
        }
        // numFmt
        if (this.nf != null)
            axisooxml.append(nf!!.getOOXML("c:"))        //need a default???: axisooxml.append("<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>");	axisooxml.append("\r\n");
        // majorTickMark
        s = this.getOption("majorTickMark")    // default= "cross"
        if (s != null) axisooxml.append("<c:majorTickMark val=\"$s\"/>")
        // minorTickMark
        s = this.getOption("minorTickMark")    // default= "cross"
        if (s != null) axisooxml.append("<c:minorTickMark val=\"$s\"/>")
        // tickLblPos
        s = this.getOption("tickLblPos")    // default= "nextTo"
        if (s != null) axisooxml.append("<c:tickLblPos val=\"$s\"/>")
        // shape properties
        if (this.spPr != null) axisooxml.append(this.spPr!!.ooxml)
        // text props
        if (this.gettxPr() != null)
            axisooxml.append(this.gettxPr()!!.ooxml)
        else if (from2003) {    // XLS->XLSX
            /*
 * label font:
 * Fontx fx= (Fontx) Chart.findRec(chartArr, Fontx.class);
   return  this.getParentChart().getWorkBook().getFont(fx.getIfnt());
*/
            var rot = 0
            val t = Chart.findRec(this.chartArr, Tick::class.java) as Tick
            if (t != null) {    // shoudn't
                rot = t.rotation.toInt()
                /**
                 * 0= no rotation (text appears left-to-right),
                 * 1=  text appears top-~~ are upright,
                 * 2= text is rotated 90 degrees counterclockwise,
                 * 3= text is rotated
                 */
                // convert BIFF8 rotation to TxPr rotation:
                when (rot) {
                    1 -> {
                    }
                    2 -> rot = -5400000
                    3 -> rot = 5400000
                }//????
                // TODO: is vert rotation from td?
                val txpr = TxPr(this.labelFont!!, rot, null)
                axisooxml.append(txpr.ooxml)
            }
            axisooxml.append("\r\n")
        }
        // crossesAx
        axisooxml.append("<c:crossAx val=\"$crossId\"/>")
        axisooxml.append("\r\n") // crosses axis ...
        // crosses -- autoZero, max, min
        if (this.getOption("crosses") != null)
            axisooxml.append("<c:crosses val=\"" + this.getOption("crosses") + "\"/>")
        axisooxml.append("\r\n")// where axis crosses it's perpendicular axis
        if (axis == "catAx" || axis == "serAx") {
            // auto
            axisooxml.append("<c:auto val=\"1\"/>\r\n")
            s = this.getOption("lblAlgn")
            if (s != null) axisooxml.append("<c:lblAlgn val=\"$s\"/>\r\n")
            s = this.getOption("lblOffset")
            if (s != null) axisooxml.append("<c:lblOffset val=\"$s\"/>\r\n")
            s = this.getOption("tickLblSkip")
            if (s != null) axisooxml.append("<c:tickLblSkip val=\"$s\"/>\r\n")
            s = this.getOption("tickMarkSkip")
            if (s != null) axisooxml.append("<c:tickMarkSkip val=\"$s\"/>\r\n")
            // TODO: noMutliLvlLbl
        } else {    // val or date
            s = this.getOption("crossBetween")
            if (s != null) axisooxml.append("<c:crossBetween val=\"$s\"/>\r\n")
            s = this.getOption("majorUnit")
            if (s != null) axisooxml.append("<c:majorUnit val=\"$s\"/>\r\n")
            s = this.getOption("minorUnit")
            if (s != null) axisooxml.append("<c:minorUnit val=\"$s\"/>\r\n")
            // TODO: dispUnit ************************************

        }
        axisooxml.append("</c:$axis>")
        axisooxml.append("\r\n")
        return axisooxml.toString()
    }


    /**
     * returns the SVG representation of the desired axis
     *
     * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
     * @return String SVG
     */
    fun getSVG(ca: ChartAxes, chartMetrics: Map<String, Double>, categories: Array<Any>): String {
        val svg = StringBuffer()

        // for all axies-- label and title fonts, rotation, tick info ...
        // Title + Label SVG
        var labelfontSVG = ""
        var titlefontSVG = ""

        try {
            labelfontSVG = this.labelFont!!.svg    // uses specific or default for chart
        } catch (e: Exception) {    // shouldn't
            labelfontSVG = "font-family='Arial' font-size='9pt' fill='" + ChartType.darkColor + "' "
        }

        try {
            titlefontSVG = linkedtd!!.getFont(this.parentChart!!.workBook!!)!!.svg
        } catch (e: NullPointerException) {
        }

        var showMinorTickMarks = false
        var showMajorTickMarks = true
        try {
            val t = Chart.findRec(chartArr, Tick::class.java) as Tick
            showMinorTickMarks = t.showMinorTicks()
            showMajorTickMarks = t.showMajorTicks()
        } catch (e: Exception) {
        }

        // BAR CHART AXES ARE SWITCHED - handle seperately for clarity; radar axes are also handled separately as are significantly different than regualr charts
        val charttype = this.parentChart!!.chartType
        if (charttype == ChartConstants.BARCHART)
            return getSVGBARCHART(ca, titlefontSVG, labelfontSVG, showMinorTickMarks, showMajorTickMarks, chartMetrics, categories)
        if (charttype == ChartConstants.RADARCHART)
            return getSVGRADARCHART(ca, titlefontSVG, labelfontSVG, chartMetrics, categories)

        var wtype = axis.toInt()
        if (wtype == ChartConstants.XAXIS && (charttype == ChartConstants.SCATTERCHART || charttype == ChartConstants.BUBBLECHART)) {    // XY Charts - X Axis is a Value Axis
            wtype = ChartConstants.XVALAXIS
        }

        when (wtype) {
            ChartConstants.XAXIS -> svg.append(drawXAxisSVG(ca, titlefontSVG, labelfontSVG, showMinorTickMarks, showMajorTickMarks, chartMetrics, categories))
            ChartConstants.YAXIS -> svg.append(drawYAxisSVG(ca, titlefontSVG, labelfontSVG, showMinorTickMarks, showMajorTickMarks, chartMetrics))
            ChartConstants.ZAXIS        // ??
            -> {
            }

            ChartConstants.XVALAXIS    // Scatter/Bubble Chart X Value Axis
            -> svg.append(drawXYValAxisSVG(ca, titlefontSVG, labelfontSVG, showMinorTickMarks, showMajorTickMarks, chartMetrics, categories))
        }
        return svg.toString()
    }

    /**
     * generate SVG for a basic (non-bar-chart, non-textual) X Axis
     *
     * @param ca
     * @param titlefontSVG
     * @param labelfontSVG
     * @param rot
     * @param showMinorTickMarks
     * @param showMajorTickMarks
     * @param chartMetrics       maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
     * @return
     */
    private fun drawXAxisSVG(ca: ChartAxes, titlefontSVG: String, labelfontSVG: String, showMinorTickMarks: Boolean, showMajorTickMarks: Boolean, chartMetrics: Map<String, Double>, categories: Array<Any>?): String {
        val svg = StringBuffer()
        // X Axis TICKS x= rectX, x2= x1, y1= recty+h, y2=y1+5 (major) x1 increments by (#Major Ticks-1)/Width (usually 4)
        //              x= rectX+17, x2= x1, y1= recty+h y2=y1+2 (minor) x1 increments by (approx 8)
        // LABELS (CATEGORIES): - do before ticks NOTE: patterns and formatting is applied .adjustCoordinates in order to account for fitting in space

        // when x axis is reversed means that categories are right to left and the y axis is on the RHS
        // when y axis is reversed means the categories are on TOP of the chart and y axis labels are reversed
        var x0: Double
        val x1: Double
        var y0: Double
        var y1: Double
        var inc: Double
        var f: java.awt.Font? = null
        val x = chartMetrics["x"]
        val y = chartMetrics["y"]
        val w = chartMetrics["w"]
        val h = chartMetrics["h"]
        val canvash = chartMetrics["canvash"]
        val yAxisReversed = ca.getMetric("yAxisReversed") as Boolean
        val xAxisReversed = ca.getMetric("xAxisReversed") as Boolean
        val xAxisRotate = ca.getMetric("xAxisRotate") as Int
        val XAXISLABELOFFSET = ca.getMetric("XAXISLABELOFFSET") as Double
        val XAXISTITLEOFFSET = ca.getMetric("XAXISTITLEOFFSET") as Double

        val labelRot = if (axis.toInt() == ChartConstants.XAXIS) xAxisRotate else 0    // TODO: handle Y axis rotation
        if (labelRot != 0) {
            // get font object so can calculate rotation point
            val lf = this.labelFont
            try {
                // get awt Font so can compute and fit category in width
                f = java.awt.Font(lf!!.fontName, lf.fontWeight, lf.fontHeightInPoints.toInt())
            } catch (e: Exception) {
            }

        }
        if (categories != null && categories.size > 0) { // shouldn't
            // Category Labels - centered within area on X Axis
            inc = w / categories.size
            svg.append(getCategoriesSVG(x, y, w, h, inc, labelRot, categories, f, labelfontSVG, yAxisReversed, xAxisReversed, XAXISLABELOFFSET))
            // TICK MARKS
            y0 = y + if (!yAxisReversed) h else 0    // ticks at bottom edge of axis unless Y axiis is reversed
            x0 = x        // start at chart x
            val rfY = if (!yAxisReversed) 1 else -1    // reverse factor :)
            val rfX = if (!xAxisReversed) 1 else -1    // reverse factor :)
            svg.append("<g>\r\n")
            inc = w / categories.size    // w/scale factor
            var minorinc = 0.0
            if (showMinorTickMarks)
                minorinc = inc / 2    // half-marks for category axis
            for (i in 0..categories.size) {
                y1 = y0 + 2 * rfY    // minor tick mark
                if (showMinorTickMarks) {
                    for (j in 0..1) {    // for categories, only option is 1/2 major
                        svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "' " + ChartType.strokeSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + y1 + "'/>\r\n")
                        x0 += minorinc
                    }
                }
                y1 = y0 + 5 * rfY    // Major tick marks
                if (showMajorTickMarks)
                    svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "'" + ChartType.strokeSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + y1 + "'/>\r\n")
                x0 += inc
            }
            // bounding edge line
            if (hasPlotAreaBorder()) {
                x0 = x + if (!xAxisReversed) w else 0
                svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "'" + ChartType.strokeSVG + " x1='" + x0 + "' y1='" + y + "' x2='" + x0 + "' y2='" + (y + h) + "'/>\r\n")
            }
            svg.append("</g>\r\n")
        }
        // X AXIS TITLE
        val titleRot = if (linkedtd != null) linkedtd!!.rotation else 0
        x0 = x + w / 2
        if (!yAxisReversed)
        // TODO: why doesn't "normal" calc work???????
            y0 = canvash - XAXISTITLEOFFSET
        else
            y0 = y - XAXISTITLEOFFSET - XAXISLABELOFFSET

        svg.append(getAxisTitleSVG(x0, y0, titlefontSVG, titleRot, "xaxistitle"))
        return svg.toString()
    }

    /**
     * generate SVG for a basic (non-bar-chart, value or numeric) Y Axis
     *
     * @param ca
     * @param titlefontSVG
     * @param labelfontSVG
     * @param rot
     * @param showMinorTickMarks
     * @param showMajorTickMarks
     * @param chartMetrics       maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
     * @return
     */
    // TODO: label rotation
    private fun drawYAxisSVG(ca: ChartAxes, titlefontSVG: String, labelfontSVG: String, showMinorTickMarks: Boolean, showMajorTickMarks: Boolean, chartMetrics: Map<String, Double>): String {
        val svg = StringBuffer()
        // Y or Value Axis- must be non-textual; values obtained in calling method
        // When Y Axis is reversed, scale is reversed and x axis labels and title are on top
        // When X Axis is reversed, Y scale/labels and title are on RHS
        var x0: Double
        var x1: Double
        var y0: Double
        var y1: Double
        val inc: Double
        // major and minor tick marks, max and min axis scales NOTE: for MOST category (usually x) axes, they are textual; the max=# of categories; min=0
        val x = chartMetrics["x"]
        val y = chartMetrics["y"]
        val w = chartMetrics["w"]
        val h = chartMetrics["h"]
        val max = chartMetrics["max"]
        val min = chartMetrics["min"]
        val minor = ca.getMetric("minor") as Double
        val major = ca.getMetric("major") as Double
        val scaleIsInteger = major == Math.floor(major)    // true if scale is in integer form rather than double (used for formatting label text)
        val titleRot = if (linkedtd != null) linkedtd!!.rotation else 0
        val xAxisReversed = ca.getMetric("xAxisReversed") as Boolean
        val yAxisReversed = ca.getMetric("yAxisReversed") as Boolean
        val xPattern = ca.getMetric("xPattern") as String
        val yPattern = ca.getMetric("yPattern") as String
        val YAXISLABELOFFSET = ca.getMetric("YAXISLABELOFFSET") as Double
        val YAXISTITLEOFFSET = ca.getMetric("YAXISTITLEOFFSET") as Double

        // Y Axis GRIDLINES: -- x1=rectX x2=rectx+rect w y1=y2  y1 starts with recty+h, decrements by= approx 27 ==rectheight/# lines *****************************
        // Major/Minor tick marks
        if (major > 0) {    // if is displaying tick marks/grid lines - usual case
            inc = h / ((max - min) / major)
            var minorinc = 0.0
            if (minor > 0)
                minorinc = inc / (major / minor)
            x0 = x
            y0 = y + h
            x1 = x + w    // entire width
            // GRIDLINES
            var lineSVG = getLineSVG(AxisLineFormat.ID_MAJOR_GRID)
            svg.append("<g>\r\n")
            run {
                var i = min
                while (i <= max) {
                    svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "'" + lineSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x1 + "' y2='" + y0 + "'/>\r\n")
                    y0 -= inc
                    i += major
                }
            }
            svg.append("</g>\r\n")
            // TICK MARKS
            svg.append("<g>\r\n")
            val rfY = if (!yAxisReversed) 1 else -1 // reverse factor :)
            val rfX = if (!xAxisReversed) 1 else -1 // reverse factor :)
            y0 = y + if (!yAxisReversed) h else 0    // starts at bottom, goes up to top except if reversed
            lineSVG = ChartType.strokeSVG
            var scale = 0
            run {
                // figure out if scale has a fractional portion; if so, determine desired scale and keep to that
                val s = major.toString()
                val z = s.indexOf(".")
                if (z != -1)
                    scale = s.length - (z + 1)
            }
            var k = 0    // axis label index
            var i = min
            while (i <= max) {
                x0 = x + if (!xAxisReversed) 0 else w
                x1 = x0 - 2 * rfX// minor ticks
                y1 = y0
                if (i < max && minor > 0) {
                    if (showMinorTickMarks) {
                        var j = 0
                        while (j < major / minor) {
                            svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "' " + lineSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x1 + "' y2='" + y0 + "'/>\r\n")
                            y0 -= minorinc
                            j++
                        }
                    }
                }
                y0 = y1
                x1 = x0 - 5 * rfX// Major tick Marks
                if (showMajorTickMarks)
                    svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "' " + lineSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x1 + "' y2='" + y0 + "'/>\r\n")
                // Y Axis Labels
                val bd = java.math.BigDecimal(i).setScale(scale, java.math.BigDecimal.ROUND_HALF_UP)    // if fractional, ensure java's floating point crap is handled corectly
                if (!xAxisReversed)
                    svg.append("<text id='yaxislabels" + k++ + "' x='" + (x0 - YLABELSSPACER_X) + "' y='" + (y1 + YLABELSPACER_Y) +
                            "' style='text-anchor: end;' direction='rtl' alignment-baseline='text-after-edge' " + labelfontSVG + ">" + CellFormatFactory.fromPatternString(yPattern).format(bd) + "</text>\r\n")
                else
                    svg.append("<text id='yaxislabels" + k++ + "' x='" + (x0 + YLABELSSPACER_X) + "' y='" + (y1 + YLABELSPACER_Y) +
                            "' style='text-anchor: start;' alignment-baseline='text-after-edge' " + labelfontSVG + ">" + CellFormatFactory.fromPatternString(yPattern).format(bd) + "</text>\r\n")
                y0 -= inc * rfY
                i += major
            }
            svg.append("</g>\r\n")
            // AXIS bounding line
            x0 = x + if (!xAxisReversed) 0 else w
            y0 = y
            svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "' " + ChartType.strokeSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + (y0 + h) + "'/>\r\n")
            // Y AXIS TITLE
            if (!xAxisReversed)
                x0 = /*YAXISTITLEOFFSET + */ 10.0    // TODO: 10 should be actual font height (when rotated 90 as is the norm)
            else
                x0 = x + w + YAXISTITLEOFFSET
            svg.append(getAxisTitleSVG(x0, y + h / 2, titlefontSVG, titleRot, "yaxistitle"))
        }
        return svg.toString()
    }

    /**
     * return the svg to display an XYValue axis (bubble, scatter)
     *
     * @param ca
     * @param titlefontSVG
     * @param labelfontSVG
     * @param showMinorTickMarks
     * @param showMajorTickMarks
     * @param chartMetrics
     * @return
     */
    private fun drawXYValAxisSVG(ca: ChartAxes, titlefontSVG: String, labelfontSVG: String, showMinorTickMarks: Boolean, showMajorTickMarks: Boolean, chartMetrics: Map<String, Double>, categories: Array<Any>?): String {
        val svg = StringBuffer()
        var x0: Double
        val x1: Double
        var y0: Double
        var y1: Double
        var minorinc = 0.0
        val inc: Double
        // Y or Value Axis- must be non-textual; values obtained in calling method
        // major and minor tick marks, max and min axis scales NOTE: for MOST category (usually x) axes, they are textual; the max=# of categories; min=0
        val x = chartMetrics["x"]
        val y = chartMetrics["y"]
        val w = chartMetrics["w"]
        val h = chartMetrics["h"]
        val max = chartMetrics["max"]
        val min = chartMetrics["min"]
        var minor = ca.getMetric("minor") as Double
        var major = ca.getMetric("major") as Double
        val scaleIsInteger = major == Math.floor(major)    // true if scale is in integer form rather than double (used for formatting label text)
        val yPattern = ca.getMetric("yPattern") as String
        var XAXISLABELOFFSET = ca.getMetric("XAXISLABELOFFSET") as Double
        val XAXISTITLEOFFSET = ca.getMetric("XAXISTITLEOFFSET") as Double

        if (categories != null && categories.size > 0) { // shouldn't
            var xmin = java.lang.Double.MAX_VALUE
            var xmax = java.lang.Double.MIN_VALUE
            var TEXTUALXAXIS = true
            for (j in categories.indices) {
                try {
                    val d = Double(categories[j].toString())
                    xmax = Math.max(xmax, d)
                    xmin = Math.min(xmin, d)
                    TEXTUALXAXIS = false        // if ANY category val is a double, assume it's a normal xyvalue axis
                } catch (e: Exception) {
                    /* keep going */
                }

            }
            if (!TEXTUALXAXIS) {
                val d = ValueRange.calcMaxMin(xmax, xmin, w)
                minor = d[0].toInt().toDouble()
                major = d[1].toInt().toDouble()
                xmax = d[2]

            } else {    // NO category values are numeric --  x axis is Textual - just use count
                major = 1.0
                minor = 0.0
                xmax = (categories.size + 1).toDouble()
                XAXISLABELOFFSET = 30.0    // TODO: calculate
            }

            // CATEGORY LABELS- centered within area on X Axis
            y0 = y + h + XAXISLABELOFFSET
            inc = w / (xmax / major)
            x0 = x
            var scale = 0
            run {
                // figure out if scale has a fractional portion; if so, determine desired scale and keep to that
                val s = major.toString()
                val z = s.indexOf(".")
                if (z != -1)
                    scale = s.length - (z + 1)
            }
            var k = 0    // axis label index
            run {
                var i = 0.0
                while (i <= xmax) {
                    if (!TEXTUALXAXIS) {
                        val bd = java.math.BigDecimal(i).setScale(scale, java.math.BigDecimal.ROUND_HALF_UP)
                        svg.append("<text id='xaxislabels" + k++ + "' x='" + x0 + "' y='" + y0 +                                            /* TODO: should really trap xyPattern */
                                "' style='text-anchor: middle;' " + labelfontSVG + ">" + CellFormatFactory.fromPatternString(yPattern).format(bd) + "</text>\r\n")
                    } else
                        svg.append("<text id='xaxislabels" + k++ + "' x='" + x0 + "' y='" + y0 +
                                "' style='text-anchor: middle;' " + labelfontSVG + ">" + CellFormatFactory.fromPatternString(yPattern).format(i) + "</text>\r\n")
                    x0 += inc
                    i += major
                }
            }

            // TICK MARKS
            y0 = h + y    // origin at y+h
            x0 = x        // start at chart x
            svg.append("<g>\r\n")
            if (minor > 0)
                minorinc = inc / minor
            var i = 0
            while (i <= xmax) {
                y1 = y0 + 2    // minor tick mark
                var j = 0
                while (j < minor) {
                    svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "' " + ChartType.strokeSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + y1 + "'/>\r\n")
                    x0 += minorinc
                    j++
                }
                y1 = y0 + 5    // Major tick marks
                svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "' " + ChartType.strokeSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + y1 + "'/>\r\n")
                if (minorinc == 0.0)
                    x0 += inc
                i += major.toInt()
            }
            svg.append("</g>\r\n")
        }
        // X AXIS TITLE
        val titleRot = if (linkedtd != null) linkedtd!!.rotation else 0
        svg.append(getAxisTitleSVG(x + w / 2, y + h + XAXISTITLEOFFSET, titlefontSVG, titleRot, "zaxistitle"))
        return svg.toString()
    }


    /**
     * Bar chart Axes are switched so handle seperately
     * basically x axis holds Y labels and title, and visa versa + gridlines go up and down (traverse y) rather than across (traverse x)
     *
     * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
     * @return String SVG
     */
    private fun getSVGBARCHART(ca: ChartAxes, titlefontSVG: String, labelFontSVG: String, showMinorTicks: Boolean, showMajorTicks: Boolean, chartMetrics: Map<String, Double>, categories: Array<Any>?): String {
        val svg = StringBuffer()

        // major and minor tick marks, max and min axis scales NOTE: for MOST category (usually x) axes, they are textual; the max=# of categories; min=0
        // X Axis/Cats in reversed order means X axis on Top, Y axis labels in reversed order (along with bars)
        var x0: Double
        var x1: Double
        var y0: Double
        var y1: Double
        var inc: Double
        var rfX = 1    // reverse factor used to reverse order + position of axes
        var rfY = 1    // ""
        var scaleIsInteger = true    // usual case
        val x = chartMetrics["x"]
        val y = chartMetrics["y"]
        val w = chartMetrics["w"]
        val h = chartMetrics["h"]
        val canvasw = chartMetrics["canvasw"]
        val canvash = chartMetrics["canvash"]
        val max = chartMetrics["max"]
        val min = chartMetrics["min"]
        val xAxisReversed = ca.getMetric("xAxisReversed") as Boolean
        val yAxisReversed = ca.getMetric("yAxisReversed") as Boolean
        val YAXISTITLEOFFSET = ca.getMetric("YAXISTITLEOFFSET") as Double
        val yPattern = ca.getMetric("yPattern") as String
        val XAXISLABELOFFSET = ca.getMetric("XAXISLABELOFFSET") as Double
        val XAXISTITLEOFFSET = ca.getMetric("XAXISTITLEOFFSET") as Double
        val major = ca.getMetric("major") as Double

        if (axis.toInt() == ChartConstants.XAXIS) {
            scaleIsInteger = major == Math.floor(major)    // true if scale is in integer form rather than double (used for formatting label text)
        }
        when (axis) {
            ChartConstants.XAXIS -> {
                // X Axis TICKS x= rectX, x2= x1, y1= recty+h, y2=y1+5 (major) x1 increments by (#Major Ticks-1)/Width (usually 4)
                //              x= rectX+17, x2= x1, y1= recty+h y2=y1+2 (minor) x1 increments by (approx 8)
                x0 = x    // start at chart x
                y0 = y + if (!xAxisReversed) h else 0    // origin at y+h (unless reversed)
                rfX = if (!xAxisReversed) 1 else -1
                rfY = if (!yAxisReversed) 1 else -1
                if (major > 0) {
                    svg.append("<g>\r\n")
                    inc = w / ((max - min) / major)    // w/scale factor
                    var scale = 0
                    run {
                        // figure out if scale has a fractional portion; if so, determine desired scale and keep to that
                        val s = major.toString()
                        val z = s.indexOf(".")
                        if (z != -1)
                            scale = s.length - (z + 1)
                    }
                    var k = 0    // axis label index
                    var i = min
                    while (i <= max) {// traverse across bottom (or top, if reversed) of x axis) incrementing x value keeping y value constant
                        y1 = y0 + 5 * rfX    // Major tick marks
                        if (showMajorTicks)
                            svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "'" + ChartType.strokeSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + y1 + "'/>\r\n")
                        // X axis labels= (Values)
                        val bd = java.math.BigDecimal(i).setScale(scale, java.math.BigDecimal.ROUND_HALF_UP)
                        if (!xAxisReversed)
                            svg.append("<text id='xaxislabels" + k++ + "' x='" + (x0 + if (!yAxisReversed) 0 else w) + "' y='" + (y0 + XAXISLABELOFFSET) + "' style='text-anchor: end;' alignment-baseline='middle' " + labelFontSVG + ">" + CellFormatFactory.fromPatternString(yPattern).format(bd) + "</text>\r\n")
                        else
                            svg.append("<text id='xaxislabels" + k++ + "' x='" + x0 + "' y='" + (y0 - 4) + "' style='text-anchor: end;' " + labelFontSVG + ">" + CellFormatFactory.fromPatternString(yPattern).format(bd) + "</text>\r\n")
                        x0 += inc * rfY
                        i += major
                    }
                    svg.append("</g>\r\n")
                }
                // AXIS bounding line
                y0 = y + if (!xAxisReversed) h else 0        // origin at y+h (unless reversed)
                x0 = x // + (!yAxisReversed?0:ci.w);		// start at chart x
                svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "' " + ChartType.strokeSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + (x0 + w) + "' y2='" + y0 + "'/>\r\n")
                // is the Y AXIS TITLE i.e. goes alongside Y AXIS
                if (!yAxisReversed)
                    x0 = YAXISTITLEOFFSET
                else
                    x0 = x + w + YAXISTITLEOFFSET
                svg.append(getAxisTitleSVG(x0, y + h / 2, titlefontSVG, 90, "xaxistitle"))
            }
            ChartConstants.YAXIS -> {
                rfX = if (!xAxisReversed) 1 else -1    // if reversed, y vals on top (x axis), cats are in reversed order (y axis)
                rfY = if (!yAxisReversed) 1 else -1    // if reversed, y vals in reverse order (x axis), cats are on RHS (y axis)
                if (categories != null && categories.size > 0) { // should!
                    inc = h / categories.size
                    x0 = x    // + (!yAxisReversed?0:ci.w);	// draw y axis tick marks + y axis labels (= categories)
                    y0 = y + if (!xAxisReversed) h else 0    // starts at bottom, goes up to top unless reversed
                    var k = 0    // axis label index
                    for (i in categories.indices) {    // traverse Y axis, spacing category labels (x is constant, y is segmented)
                        x1 = x0 - 5 * rfY// Major tick Marks
                        if (showMajorTicks)
                            svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "' " + ChartType.strokeSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x1 + "' y2='" + y0 + "'/>\r\n")
                        // Y Axis Labels = categories
                        y1 = y0 - inc * rfX + inc * .5 * rfX.toDouble()
                        if (!yAxisReversed)
                            svg.append("<text id='yaxislabels" + k++ + "' x='" + (x0 - 8) + "' y='" + y1 +
                                    "' style='text-anchor: end;' direction='rtl' dominant-baseline='text-before-edge' " + labelFontSVG + ">" + categories[i].toString() + "</text>\r\n")
                        else
                            svg.append("<text id='yaxislabels" + k++ + "' x='" + (x0 + w + 8.0) + "' y='" + (y1 + 4) +
                                    "' style='text-anchor: start;' alignment-baseline='text-after-edge' " + labelFontSVG + ">" + categories[i].toString() + "</text>\r\n")
                        y0 -= inc * rfX
                    }
                    // show gridlines (top to bottom)
                    val lineSVG = getLineSVG(AxisLineFormat.ID_MAJOR_GRID)
                    if (lineSVG != "") {
                        y0 = y
                        x0 = x    // start at chart x
                        inc = w / ((max - min) / major)    // w/scale factor
                        // 1st line is axis line
                        svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "' " + getLineSVG(AxisLineFormat.ID_AXIS_LINE) + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + (y0 + h) + "'/>\r\n")
                        // rest are grid lines
                        var i = min
                        while (i < max) {
                            x0 += inc
                            svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "' " + lineSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + (y0 + h) + "'/>\r\n")
                            i += major
                        }
                    }
                    // is the X AXIS TITLE i.e. GOES ON THE X AXIS
                    y0 = if (!xAxisReversed) canvash - XAXISTITLEOFFSET else y - XAXISTITLEOFFSET
                    val titleRot = if (linkedtd != null) linkedtd!!.rotation else 0
                    svg.append(getAxisTitleSVG(x + w / 2, y0, titlefontSVG, titleRot, "yaxistitle"))
                }
            }
        }
        return svg.toString()
    }

    /**
     * returns the SVG representation of this Radar Chart Axes --
     * <br></br>Radar Chart Axis look like a spider web
     * <br></br>Note that these coordinates, like all axis scales, match the coordinates and calculations in Rader.getSVG
     *
     * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
     * @return String SVG
     */
    private fun getSVGRADARCHART(ca: ChartAxes, titleFontSVG: String, labelFontSVG: String, chartMetrics: Map<String, Double>, categories: Array<Any>?): String {
        val x = chartMetrics["x"]
        val y = chartMetrics["y"]
        val w = chartMetrics["w"]
        val h = chartMetrics["h"]
        val max = chartMetrics["max"]
        val min = chartMetrics["min"]
        var major = ca.getMetric("major") as Double
        val svg = StringBuffer()
        if (axis.toInt() == ChartConstants.YAXIS) {
            svg.append("<g>\r\n")
            if (categories != null && categories.size > 0) { // shouldn't
                major = major * 2    // appears the scale calc doesn't follow other charts ...
                val n = categories.size.toDouble()
                val centerx = w / 2 + x
                val centery = h / 2 + y
                val percentage = 1 / n        // divide into equal sections
                var radius = Math.min(w, h) / 2.3    // should take up almost entire w/h of chart
                val radiusinc = radius / (max / major)
                var lastx = centerx
                var lasty = centery - radius    // again, start straight up
                var k = 0    // axis label index
                var j = min
                while (j <= max) {    // each major unit is a concentric line
                    var angle = 90.0            // starts straight up
                    var i = 0
                    while (i <= n) {        // each category is a radial line; <= n so can complete the spider web line path
                        // get next point on circumference
                        val x1 = centerx + radius * Math.cos(Math.toRadians(angle))
                        val y1 = centery - radius * Math.sin(Math.toRadians(angle))
                        svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "'" + ChartType.strokeSVG + " x1='" + lastx + "' y1='" + lasty + "' x2='" + x1 + "' y2='" + y1 + "'/>\r\n")
                        if (j == 0.0 && i < n) {    // print radial lines & labels at very end of chart/top of each line  -- only do for 1st round
                            svg.append("<line fill='none' fill-opacity='" + ChartType.fillOpacity + "' " + ChartType.strokeSVG + " x1='" + centerx + "' y1='" + centery + "' x2='" + x1 + "' y2='" + y1 + "'/>\r\n")
                            val labelx1 = centerx + (radius + 10) * Math.cos(Math.toRadians(angle))
                            val labely1 = centery - (radius + 10) * Math.sin(Math.toRadians(angle))
                            svg.append("<text id='xaxislabels" + k++ + "' x='" + labelx1 + "' y='" + labely1 + "' style='text-anchor: middle;' " + labelFontSVG + ">" + categories[i].toString() + "</text>\r\n")
                        }
                        // next angle
                        angle -= percentage * 360
                        lastx = x1
                        lasty = y1
                        i++
                    }
                    radius -= radiusinc
                    j += major
                }
            }
            svg.append("</g>\r\n")
        }
        return svg.toString()
    }


    /**
     * return the svg necessary to display categories along an x axis
     *
     * @param x
     * @param y
     * @param w
     * @param h
     * @param inc
     * @param labelRot
     * @param categories
     * @param f
     * @param labelfontSVG
     * @param yAxisReversed
     * @param xAxisReversed
     * @param XAXISLABELOFFSET
     * @return
     */
    private fun getCategoriesSVG(x: Double, y: Double, w: Double, h: Double, inc: Double, labelRot: Int, categories: Array<Any>, f: java.awt.Font?, labelfontSVG: String, yAxisReversed: Boolean, xAxisReversed: Boolean, XAXISLABELOFFSET: Double): String {
        // Category Labels - centered within area on X Axis
        val svg = StringBuffer()
        var x0: Double
        val x1: Double
        var y0: Double
        val y1: Double
        val k = labelfontSVG.indexOf("font-size=") + 11
        val fh = java.lang.Double.parseDouble(labelfontSVG.substring(k, labelfontSVG.indexOf("pt")))    // approximate height of a line of labels
        y0 = y + if (!yAxisReversed) h + XAXISLABELOFFSET / 3 else -XAXISLABELOFFSET    // draw on bottom edge of axis unless Y axis is reversed
        var m = 0    // axis label index
        for (i in categories.indices) {
            if (!xAxisReversed)
                x0 = x + inc * i + inc / 2
            else
            // reversed:  category labels start from LHS and go to RHS
                x0 = x + w - inc * i - inc / 2
            if (labelRot != 0) {
                var len = StringTool.getApproximateStringWidthLB(f, CellFormatFactory.fromPatternString(null).format(categories[i]))
                if (labelRot == 45)
                    len = Math.ceil(len * Math.cos(Math.toRadians(labelRot.toDouble()))).toInt().toDouble()
                val offset = (len / 2).toInt() + 5
                y0 = y + if (!yAxisReversed) h + offset else -offset
                if (labelRot == 45)
                    x0 += inc / 2

            }
            // handle multiple lines in X axis labels - must do "by hand"
            val s = categories[i].toString().split("\n".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
            svg.append("<text id='xaxislabels" + m++ + "' x='" + x0 + "' y='" + y0 +
                    (if (labelRot == 0) "" else "' transform='rotate($labelRot, $x0 , $y0)") +
                    "' style='text-anchor: middle;' alignment-baseline='text-after-edge' " + labelfontSVG + ">")
            for (z in s.indices)
                svg.append("<tspan x='" + x0 + "' dy='" + fh * 1.4 + "'>" + s[z] + "</tspan>\r\n")
            svg.append("</text>\r\n")
        }
        return svg.toString()
    }

    /**
     * return the svg necessary to display an axis title
     *
     * @param x
     * @param y
     * @param titlefontSVG
     * @param titleRot     0, 45 or -90 degrees
     * @param scriptTitle  either "xaxistitle", "yaxistitle" or "zaxistitle"
     * @return
     */
    private fun getAxisTitleSVG(x: Double, y: Double, titlefontSVG: String, titleRot: Int, scriptTitle: String): String {
        val svg = StringBuffer()
        svg.append("<g>\r\n")
        svg.append("<text " + GenericChartObject.getScript(scriptTitle) + " x='" + x + "' y='" + y +
                (if (titleRot == 0) "" else "' transform='rotate(-$titleRot, $x ,$y)") +
                "' style='text-anchor: middle;' " + titlefontSVG + ">" + this.title + "</text>\r\n")
        svg.append("</g>\r\n")
        return svg.toString()
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -8592219101790307789L
        // Axis placement
        val INVISIBLE = 0
        val LOW = 1
        val HIGH = 2
        val NEXTTO = 3

        fun getPrototype(wType: Int): XLSRecord {
            val a = Axis()
            a.opcode = XLSConstants.AXIS
            a.data = a.PROTOTYPE_BYTES
            //        if (wType!=XVALAXIS)
            a.setAxis(wType)
            //        else
            //        	a.setAxis(XAXIS);
            // also add the associated records
            when (wType) {
                ChartConstants.XAXIS -> {
                    a.addChartRecord(CatserRange.prototype)
                    a.addChartRecord(Axcent.prototype)
                    a.addChartRecord(Tick.prototype)
                }
                ChartConstants.YAXIS, ChartConstants.XVALAXIS -> {
                    a.addChartRecord(ValueRange.prototype)
                    a.addChartRecord(Tick.prototype)
                    val alf = AxisLineFormat.prototype as AxisLineFormat?
                    alf!!.setId(AxisLineFormat.ID_MAJOR_GRID)    // default has major gridlines
                    a.addChartRecord(alf)
                    a.addChartRecord(LineFormat.prototype)
                }
                ChartConstants.ZAXIS -> {
                    // KSC: TODO: Set CatserRange options correctly when get def!!! **********
                    val c = CatserRange.prototype as CatserRange?
                    a.addChartRecord(c)
                    a.addChartRecord(Tick.prototype)    // TODO: Tick should have
                }
            }
            return a
        }

        internal var YLABELSSPACER_X = 10
        internal var YLABELSPACER_Y = 4
    }

}

internal class Scaling : OOXMLElement {
    private var logBase: String? = null
    private var max: String? = null
    private var min: String? = null
    private var orientation: String? = null

    override// sequence:  logBase, orientation, max, min
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<c:scaling")
            if (logBase != null) ooxml.append("<c:logBase val=\"$logBase\"/>")
            if (orientation != null) ooxml.append("<c:orientation val=\"$orientation\"/>")
            if (max != null) ooxml.append("<c:max val=\"$max\"/>")
            if (min != null) ooxml.append("<c:min val=\"$min\"/>")
            ooxml.append("</scaling>\r\n")
            return ooxml.toString()
        }

    constructor() {    // no-param constructor, set up common defaults
    }

    constructor(logBase: String, max: String, min: String, orientation: String) {
        this.logBase = logBase
        this.max = max
        this.min = min
        this.orientation = orientation
    }

    constructor(sc: Scaling) {
        this.logBase = sc.logBase
        this.max = sc.max
        this.min = sc.min
        this.orientation = sc.orientation
    }

    /**
     * return the specific scaling option
     *
     * @param op
     * @return
     */
    fun getOption(op: String): String? {
        if (op == "logBase")
            return logBase
        else if (op == "max")
            return max
        else if (op == "min")
            return min
        else if (op == "orientation")
            return orientation
        return null
    }

    override fun cloneElement(): OOXMLElement {
        return Scaling(this)
    }

    companion object {

        /**
         * parse Axis OOXML element (catAx, valAx, serAx or dateAx)
         *
         * @param xpp     XmlPullParser
         * @param lastTag element stack
         * @return spPr object
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<*>): OOXMLElement {
            var logBase: String? = null
            var max: String? = null
            var min: String? = null
            var orientation: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "scaling") {
                        } else if (tnm == "logBase") {
                            logBase = xpp.getAttributeValue(0)
                        } else if (tnm == "max") {
                            max = xpp.getAttributeValue(0)
                        } else if (tnm == "min") {
                            min = xpp.getAttributeValue(0)
                        } else if (tnm == "orientation") {
                            orientation = xpp.getAttributeValue(0)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "scaling") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("scaling.parseOOXML: $e")
            }

            return Scaling(logBase, max, min, orientation)
        }
    }
}