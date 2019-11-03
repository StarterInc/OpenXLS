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
import io.starter.OpenXLS.ChartHandle.ChartOptions
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.OOXML.OOXMLConstants
import io.starter.formats.XLS.*
import io.starter.formats.XLS.charts.ChartType.Companion
import io.starter.formats.cellformat.CellFormatFactory
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import org.json.JSONArray
import org.json.JSONException
import org.json.JSONObject
import org.xmlpull.v1.XmlPullParser

import java.io.BufferedWriter
import java.io.FileWriter
import java.io.IOException
import java.io.Serializable
import java.util.ArrayList
import java.util.EnumSet
import java.util.HashMap

/**
 * This abstract class defines a Chart Group for a BIFF Chart.  An Excel Chart may have up to 4 (2003 and previous versions) or 9 (post-2003) charts within one chart.
 * The 0th chart type in the Chart Group is the default chart.
 * <br></br>
 * NOTES:
 * <br></br>
 * CRT = ChartFormat Begin (Bar / Line / (BopPop [BopPopCustom]) / Pie / Area / Scatter / Radar / RadarArea / Surf) CrtLink [SeriesList] [Chart3d] [LD] [2DROPBAR] *4(CrtLine LineFormat) *2DFTTEXT [DataLabExtContents] [SS] *4SHAPEPROPS End
 *
 *
 * MUST CHANGE ChartObject if change chart type
 * MUST ADD/REMOVE ChartObject if add/remove multiple charts
 * axisparent  --> when add a new chart a new crt is added  axes + titles/labels, number format, etc
 *
 *
 * REFACTOR:  TO FINISH
 * showDataTable -- Finish, TODO
 * NOTE: The SeriesList record specifies the series of the chart. This record MUST NOT exist in the first chart group in the chart sheet substream.
 * This record MUST exist when not in the first chart group in the chart sheet substream
 *
 *
 * NOTE:
 * The Chart3d record specifies that the plot area, axis group, and chart group are rendered in a 3-D scene, rather than a 2-D scene, and specifies properties of the 3-D scene. If this record exists in the chart sheet substream, the chart sheet substream  MUST have exactly one chart group. This record MUST NOT exist in a bar of pie, bubble, doughnut, filled radar, pie of pie, radar, or scatter chart group.
 *
 *
 * NOTE: legends only in 1st chart group
 */
abstract class ChartType : ChartConstants, Serializable {
    protected var chartobj: GenericChartObject
    /**
     * return the data legend for this chart
     *
     * @return
     */
    var dataLegend: Legend? = null
        protected set
    var cf: ChartFormat? = null
    //    protected ChartSeries chartseries= new ChartSeries();
    @Transient
    protected var wb: WorkBook? = null
    protected var defaultShape = 0                            // controls default bar shape for all bars in the chart; used when adding or setting series

    val parentChart: Chart?
        get() = chartobj.parentChart

    /**
     * return data label options for each series as an int array
     * <br></br>each can be one or more of:
     * <br></br>VALUELABEL= 0x1;
     * <br></br>VALUEPERCENT= 0x2;
     * <br></br>CATEGORYPERCENT= 0x4;
     * <br></br>SMOOTHEDLINE= 0x8;
     * <br></br>CATEGORYLABEL= 0x10;
     * <br></br>BUBBLELABEL= 0x20;
     * <br></br>SERIESLABEL= 0x40;
     *
     * @return int array
     * @see AttachedLabel
     */
    protected// data label options, if any, per series
    val dataLabelInts: IntArray
        get() = chartobj.parentChart!!.getDataLabelsPerSeries(cf!!.dataLabelsInt)

    /**
     * return the default data label setting for the chart, if any
     * <br></br>NOTE: each series can override the default data label for the chart
     * <br></br>can be one or more of:
     * <br></br>VALUELABEL= 0x1;
     * <br></br>VALUEPERCENT= 0x2;
     * <br></br>CATEGORYPERCENT= 0x4;
     * <br></br>SMOOTHEDLINE= 0x8;
     * <br></br>CATEGORYLABEL= 0x10;
     * <br></br>BUBBLELABEL= 0x20;
     * <br></br>SERIESLABEL= 0x40;
     *
     * @return int default data label for chart
     */
    val dataLabel: Int
        get() = cf!!.dataLabelsInt

    /**
     * return an array of the type of markers for each series:
     * <br></br>0 = no marker
     * <br></br>1 = square
     * <br></br>2 = diamond
     * <br></br>3 = triangle
     * <br></br>4 = X
     * <br></br>5 = star
     * <br></br>6 = Dow-Jones
     * <br></br>7 = standard deviation
     * <br></br>8 = circle
     * <br></br>9 = plus sign
     */
    val markerFormats: IntArray
        get() {
            val mf = cf!!.markerFormat
            val markers = IntArray(parentChart!!.getAllSeries(parentChart!!.getChartOrder(this)).size)
            if (mf > 0) {
                for (i in markers.indices) {
                    markers[i] = mf
                }
            }
            return markers
        }

    /**
     * returns true if this chart has smoothed lines
     *
     * @return
     */
    var hasSmoothLines: Boolean
        get() = cf!!.hasSmoothLines
        set(b) {
            cf!!.hasSmoothLines = b
        }


    /**
     * returns the chart type for the default chart
     */
    val chartType: Int
        get() = chartobj.chartType

    val svg: String?
        get() = null

    val json: String?
        get() = null

    val optionsJSON: JSONObject?
        get() = null

    /**
     * return Type JSON for generic chart types
     *
     * @return
     * @throws JSONException
     */
    val typeJSON: JSONObject
        @Throws(JSONException::class)
        get() {
            val typeJSON = JSONObject()
            typeJSON.put("type", "Default")
            return typeJSON
        }

    /**
     * return the Data Labels chosen for this chart, if any
     * can be one or more of:
     * <br></br>Value
     * <br></br>ValuePerecentage
     * <br></br>CategoryPercentage
     * <br></br>CategoryLabel
     * <br></br>BubbleLabel
     * <br></br>SeriesLabel
     * or an empty string if no data labels are chosen for the chart
     *
     * @return
     */
    val dataLabels: String?
        get() = cf!!.dataLabels

    /**
     * return data label options as an int
     * <br></br>can be one or more of:
     * <br></br>VALUELABEL= 0x1;
     * <br></br>VALUEPERCENT= 0x2;
     * <br></br>CATEGORYPERCENT= 0x4;
     * <br></br>SMOOTHEDLINE= 0x8;
     * <br></br>CATEGORYLABEL= 0x10;
     * <br></br>BUBBLELABEL= 0x20;
     * <br></br>SERIESLABEL= 0x40;
     *
     * @return a combination of data label options above or 0 if none
     * @see AttachedLabel
     */
    val dataLabelsInt: Int
        get() = cf!!.dataLabelsInt

    /**
     * returns the bar shape for a column or bar type chart
     * can be one of:
     * <br></br>ChartConstants.SHAPECOLUMN	default
     * <br></br>ChartConstants.SHAPECONEd
     * <br></br>ChartConstants.SHAPECONETOMAX
     * <br></br>ChartConstants.SHAPECYLINDER
     * <br></br>ChartConstants.SHAPEPYRAMID
     * <br></br>ChartConstants.SHAPEPYRAMIDTOMAX
     *
     * @return int bar shape
     */
    val barShape: Int
        get() = ChartConstants.SHAPEDEFAULT

    /**
     * returns type of marker for this chart, if any
     * <br></br>0 = no marker
     * <br></br>1 = square
     * <br></br>2 = diamond
     * <br></br>3 = triangle
     * <br></br>4 = X
     * <br></br>5 = star
     * <br></br>6 = Dow-Jones
     * <br></br>7 = standard deviation
     * <br></br>8 = circle
     * <br></br>9 = plus sign
     *
     * @return int marker type
     */
    val markerFormat: Int
        get() = cf!!.markerFormat

    /**
     * returns true if this chart has lines (see Scatter, Line Charts amongst others)
     *
     * @return
     */
    val hasLines: Boolean
        get() = cf!!.hasLines


    /**
     * returns true if this chart has drop lines
     *
     * @return
     */
    val hasDropLines: Boolean
        get() = cf!!.hasDropLines

    /**
     * return chart-type-specific options in XML form
     *
     * @param nChart 0=default, 1-9= overlay charts
     * @return String XML
     */
    // chart-type-specific
    // governs threed settings and other misc. options
    val chartOptionsXML: String
        get() {
            val sb = StringBuffer()
            sb.append(chartobj.optionsXML)
            sb.append(cf!!.chartOptionsXML)
            return sb.toString()
        }


    /**
     * @return truth of "Chart is Three-D"
     */
    val isThreeD: Boolean
        get() = cf!!.isThreeD(chartobj.chartType)

    /**
     * return ThreeD settings for this chart in XML form
     *
     * @return String XML
     */
    val threeDXML: String
        get() = cf!!.threeDXML


    /**
     * @return truth of "Chart is Stacked"
     */
    var isStacked: Boolean
        get() = chartobj.isStacked
        set(isstacked) {
            chartobj.isStacked = isstacked
        }

    /**
     * return truth of "Chart is 100% Stacked"
     *
     * @return
     */
    val is100PercentStacked: Boolean
        get() = chartobj.is100Percent

    /**
     * @return truth of "Chart is Clustered"  (Bar/Col only)
     */
    val isClustered: Boolean
        get() = false

    /**
     * returns the SVG for the font style of this object
     */
    protected val fontSVG: String
        get() = getFontSVG(-1)

    /**
     * convert the default shape flag to a user-friendly (OOXML-compliant) String
     *
     * @param shape int
     * @return
     */
    val shape: String?
        get() {
            when (defaultShape) {
                0 -> return "box"
                1 -> return "cylinder"
                256 -> return "pyramid"
                257 -> return "cone"
                512 -> return "pyramidToMax"
                513 -> return "coneToMax"
            }
            return null
        }

    /**
     * returns an int representing the space between points in a 3d area, bar or line chart, or 0 if not 3d
     *
     * @return
     */
    val gapDepth: Int
        get() = cf!!.gapDepth

    /**
     * return the SeriesList record for this chart object
     * The SeriesList record maps the series for the chart.
     *
     * @return
     */
    val seriesList: SeriesList
        get() = Chart.findRec(cf!!.chartArr, SeriesList::class.java) as SeriesList

    constructor() {}

    constructor(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) {
        chartobj = charttype    // record that defines the chart type (Bar, Area, Line ...)
        this.wb = wb
        this.cf = cf
    }

    fun setOptions(options: EnumSet<ChartOptions>) {
        // FYI: The CrtLine (section 2.4.68) LineFormat (section 2.4.156) record pairs and the sequences of records that conform to the SHAPEPROPS rule (section 2.1.7.20.1)
        // specify the drop lines, high-low lines, series lines, and leader lines for the chart (section 2.2.3.3).
        // NO 3d record for: bar of pie, bubble, doughnut, filled radar, pie of pie, radar, or scatter chart group.
        // Has Ser (Z) axis:  Surface, fStacked==0 & Line, Area, fStacked==0 && fClustered==0 && Bar (Col)  (Must also have ThreeD record)
        // 2 Value Axes:  Scatter, Bubble
        // NO Axes:  Pie, Doughnut, PieOfPie, BarOfPie


        val ca = this.parentChart!!.axes
        val chartType = this.chartType
        if (options.contains(ChartOptions.STACKED))
        // bar/col types, line, area
            chartobj.isStacked = true

        if (options.contains(ChartOptions.PERCENTSTACKED))
        // bar/col types, line, area
            chartobj.is100Percent = true

        if (options.contains(ChartOptions.CLUSTERED))
        // bar/col only
            cf!!.setIsClustered(true)

        if (options.contains(ChartOptions.SERLINES))
        // bar, line, stock
            cf!!.addChartLines(ChartLine.TYPE_SERIESLINE.toInt())

        if (options.contains(ChartOptions.HILOWLINES))
        // bar, OfPie
            cf!!.addChartLines(ChartLine.TYPE_HILOWLINE.toInt())

        if (options.contains(ChartOptions.DROPLINES))
        // Surface chart
            cf!!.addChartLines(ChartLine.TYPE_DROPLINE.toInt())

        if (options.contains(ChartOptions.UPDOWNBARS))
        // line, area,stock
            cf!!.addUpDownBars()

        if (options.contains(ChartOptions.HASLINES))
        // line, scatter ...
            cf!!.setHasLines()

        if (options.contains(ChartOptions.SMOOTHLINES))
        // line, scatter, radar
            cf!!.hasSmoothLines = true
        // HANDLE FILLED for radar
        // HANDLE bubble 3d
        // HANDLE bar shapes ***

        cf!!.setChartObject(chartobj)
        var use3Ddefaults = true            // init 3D record with default values for specific chart type
        var threeD = cf!!.getThreeDRec(false)
        if (threeD == null) {
            if (options.contains(ChartOptions.THREED) || chartType == ChartConstants.SURFACECHART) {    // surface charts ALWAYS have a 3 record as does pyramid, cone and cylinder charts
                if (chartType != ChartConstants.BUBBLECHART && chartType != ChartConstants.SCATTERCHART)
                // supposed to be also donught, radar as well ...
                    threeD = this.initThreeD(chartType)
                else if (chartType == ChartConstants.BUBBLECHART) { // scatter charts have no 3d option
                    cf!!.setHas3DBubbles(true)
                }
            }
        } else
        // 3D record already set (via OOXML) - do not use defaults
            use3Ddefaults = false
        when (chartType) {
            ChartConstants.BARCHART, ChartConstants.COLCHART -> {
                if (use3Ddefaults && threeD != null) {
                    threeD.setChartOption("AnRot", "20")
                    threeD.setChartOption("AnElev", "15")
                    threeD.setChartOption("TwoDWalls", "true")
                    threeD.setChartOption("ThreeDScaling", "false")
                    threeD.setChartOption("Cluster", "false")
                    threeD.setChartOption("PcDepth", "100")
                    threeD.setChartOption("PcDist", "30")
                    threeD.setChartOption("PcGap", "150")
                    threeD.setChartOption("PcHeight", "72")    // ??????
                    threeD.setChartOption("Perspective", "false")
                    if (options.contains(ChartOptions.CLUSTERED)) {
                        threeD.setChartOption("Cluster", "true")
                        threeD.setChartOption("PcHeight", "62")    // ??????
                        if (chartType == ChartConstants.COLCHART)
                            threeD.setChartOption("Perspective", "true")
                        else {    // bar chart
                            threeD.setChartOption("Perspective", "false")
                            threeD.setChartOption("ThreeDScaling", "true")
                            threeD.setChartOption("PcHeight", "150")    // ??????
                        }
                    }
                }
                ca!!.setChartOption(ChartConstants.XAXIS, "AddArea", "true")
                ca.setChartOption(ChartConstants.YAXIS, "AddArea", "true")
                if (!isStacked && !isClustered) {
                    ca.createAxis(ChartConstants.ZAXIS)
                    ca.setChartOption(ChartConstants.ZAXIS, "AddArea", "true")
                }
            }
            ChartConstants.LINECHART -> if (options.contains(ChartOptions.THREED)) {
                ca!!.setChartOption(ChartConstants.XAXIS, "AddArea", "true")
                ca.setChartOption(ChartConstants.YAXIS, "AddArea", "true")
                if (!isStacked) {
                    ca.createAxis(ChartConstants.ZAXIS)
                    ca.setChartOption(ChartConstants.ZAXIS, "AddArea", "true")
                }
            }
            ChartConstants.STOCKCHART -> {
                cf!!.addChartLines(ChartLine.TYPE_HILOWLINE.toInt())
                cf!!.setMarkers(0)
            }
            ChartConstants.SCATTERCHART -> if (!options.contains(ChartOptions.HASLINES)) {
                cf!!.setHasLines(5)    // no line style -- doesn't appear to work
            }
            ChartConstants.AREACHART -> {
                if (use3Ddefaults && threeD != null) {
                    threeD.setChartOption("AnRot", "20")
                    threeD.setChartOption("TwoDWalls", "true")
                    threeD.setChartOption("ThreeDScaling", "false")
                    threeD.setChartOption("Perspective", "true")
                }
                this.setChartOption("Percentage", "25")
                this.setChartOption("SmoothedLine", "true")
                ca!!.setChartOption(ChartConstants.XAXIS, "CrossBetween", "false")
                if (options.contains(ChartOptions.THREED)) {
                    ca.setChartOption(ChartConstants.XAXIS, "AddArea", "true")
                    ca.setChartOption(ChartConstants.YAXIS, "AddArea", "true")
                    if (!isStacked) {
                        ca.createAxis(ChartConstants.ZAXIS)
                        ca.setChartOption(ChartConstants.ZAXIS, "AddArea", "true")
                    }
                }
            }
            ChartConstants.BUBBLECHART -> if (options.contains(ChartOptions.THREED)) {
                this.setChartOption("Percentage", "25")
                this.setChartOption("SmoothedLine", "true")
                this.setChartOption("ThreeDBubbles", "true")
            }
            ChartConstants.PIECHART // Pie
            -> {
                cf!!.varyColor = true // Should be true for all pie charts ...
                if (options.contains(ChartOptions.EXPLODED)) {
                    this.setChartOption("SmoothedLine", "true")
                    this.setChartOption("Percentage", "25")
                }
                if (use3Ddefaults && options.contains(ChartOptions.THREED))
                    this.setChartOption("AnRot", "236")
                ca!!.removeAxes()
            }
            ChartConstants.DOUGHNUTCHART -> {
                if (options.contains(ChartOptions.EXPLODED)) {
                    this.setChartOption("SmoothedLine", "true")
                    this.setChartOption("Percentage", "25")
                }
                cf!!.varyColor = true // Should be true for all pie charts ...
                ca!!.removeAxes()
            }
            ChartConstants.SURFACECHART    // NOTE: For Surface charts, non-threeD==Contour
            -> {
                if (use3Ddefaults && threeD != null) {    // shouldn't
                    threeD.setChartOption("Cluster", "false")
                    threeD.setChartOption("TwoDWalls", "true")
                    threeD.setChartOption("ThreeDScaling", "true")
                    threeD.setChartOption("Perspective", "true")
                }
                ca!!.setChartOption(ChartConstants.XAXIS, "AddArea", "true")
                ca.setChartOption(ChartConstants.XAXIS, "AreaFg", "8")
                ca.setChartOption(ChartConstants.XAXIS, "AreaBg", "78")
                ca.setChartOption(ChartConstants.YAXIS, "AddArea", "true")
                ca.setChartOption(ChartConstants.YAXIS, "AreaFg", "22")
                ca.setChartOption(ChartConstants.YAXIS, "AreaBg", "78")
                ca.createAxis(ChartConstants.ZAXIS)
                if (use3Ddefaults && options.contains(ChartOptions.THREED)) {    // "regular" 3d surface
                    threeD!!.setChartOption("AnElev", "15")
                    threeD.setChartOption("AnRot", "20")
                    threeD.setChartOption("PcDepth", "100")
                    threeD.setChartOption("PcDist", "30")
                    threeD.setChartOption("PcGap", "150")
                    threeD.setChartOption("PcHeight", "50")    // ??????
                } else if (use3Ddefaults) {                                        // contour (non-3d)
                    threeD!!.setChartOption("AnElev", "90")
                    threeD.setChartOption("AnRot", "0")
                    threeD.setChartOption("PcDepth", "100")
                    threeD.setChartOption("PcDist", "0")
                    threeD.setChartOption("PcGap", "150")
                    threeD.setChartOption("PcHeight", "50")    // ??????
                }
                if (options.contains(ChartOptions.WIREFRAME))
                    (chartobj as Surface).isWireframe = true
                if (!options.contains(ChartOptions.WIREFRAME) && options.contains(ChartOptions.THREED)) {
                    chartobj.setChartOption("ColorFill", "true")
                    chartobj.setChartOption("Shading", "true")
                } else if (options.contains(ChartOptions.WIREFRAME)) {    // conotur (flat, non-3d) wireframe
                    chartobj.setChartOption("Shading", "true")
                } else
                // contour filled (i.e. plain Surface)
                    chartobj.setChartOption("ColorFill", "true")
            }
            ChartConstants.RADARCHART -> if (options.contains(ChartOptions.FILLED))
                (this as RadarChart).isFilled = true
            ChartConstants.PYRAMIDCHART, ChartConstants.CONECHART, ChartConstants.CYLINDERCHART, ChartConstants.PYRAMIDBARCHART, ChartConstants.CONEBARCHART, ChartConstants.CYLINDERBARCHART -> {
                // Shaped Bar/Col charts are all 3d
                if (threeD == null) threeD = this.initThreeD(chartType)
                if (use3Ddefaults) {
                    threeD.setChartOption("AnElev", "15")
                    threeD.setChartOption("AnRot", "20")
                    threeD.setChartOption("Cluster", "true") // only for regular pyramid charts; stacked, etc. will alter
                    threeD.setChartOption("TwoDWalls", "true")
                    if (!options.contains(ChartOptions.THREED)) {
                        threeD.setChartOption("ThreeDScaling", "false")
                        threeD.setChartOption("Cluster", "true")
                    }
                    threeD.setChartOption("Perspective", "false")
                    threeD.setChartOption("PcDepth", "100")
                    threeD.setChartOption("PcDist", "30")
                    threeD.setChartOption("PcGap", "150")
                    threeD.setChartOption("PcHeight", "52")    // ??????
                }
                if (chartType == ChartConstants.PYRAMIDCHART || chartType == ChartConstants.PYRAMIDBARCHART)
                    cf!!.barShape = ChartConstants.SHAPEPYRAMID
                else if (chartType == ChartConstants.CONECHART || chartType == ChartConstants.CONEBARCHART)
                    cf!!.barShape = ChartConstants.SHAPECONE
                else if (chartType == ChartConstants.CYLINDERCHART || chartType == ChartConstants.CYLINDERBARCHART)
                    cf!!.barShape = ChartConstants.SHAPECYLINDER
            }
        }
    }

    /**
     * specifies whether
     * the color for each data point and the color and type for each data marker
     * vary
     *
     * @param b
     */
    protected fun setVaryColor(b: Boolean) {
        cf!!.varyColor = b
    }

    fun addLegend(l: Legend) {
        dataLegend = l
    }

    /**
     * gets the chart-type specific OOXML representation (representing child element of plotArea element)
     *
     * @return
     */
    open fun getOOXML(catAxisId: String, valAxisId: String, serAxisId: String): StringBuffer? {
        return null
    }

    @Throws(JSONException::class)
    fun getJSON(s: ChartSeries, wbh: WorkBookHandle, minMax: Array<Double>): JSONObject {
        val chartObjectJSON = JSONObject()

        // Type JSON
        chartObjectJSON.put("type", this.typeJSON)

        // Deal with Series
        var yMax = 0.0
        var yMin = 0.0
        var nSeries = 0
        val seriesJSON = JSONArray()
        val seriesCOLORS = JSONArray()
        try {
            val series = s.seriesRanges
            val scolors = s.seriesBarColors
            for (i in series!!.indices) {
                val seriesvals = CellRange.getValuesAsJSON(series[i].toString(), wbh)
                // must trap min and max for axis tick and units
                nSeries = Math.max(nSeries, seriesvals.length())
                for (j in 0 until seriesvals.length()) {
                    try {
                        yMax = Math.max(yMax, seriesvals.getDouble(j))
                        yMin = Math.min(yMin, seriesvals.getDouble(j))
                    } catch (n: NumberFormatException) {
                    }

                }
                seriesJSON.put(seriesvals)
                seriesCOLORS.put(scolors!![i])
            }
            chartObjectJSON.put("Series", seriesJSON)
            chartObjectJSON.put("SeriesFills", seriesCOLORS)
        } catch (je: JSONException) {
            // TODO: Log error
        }

        minMax[0] = yMin
        minMax[1] = yMax
        minMax[2] = nSeries
        return chartObjectJSON
    }

    /**
     * returns SVG to represent the actual chart object (BAR, LINE, AREA, etc.)
     *
     * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
     * @param axisMetrics  maps specific axis options such as xAxisReversed, xPattern ...
     * @param s            ChartSeries - holds legends, categories, seriesdata ...
     * @return String svg
     */
    open fun getSVG(chartMetrics: HashMap<String, Double>, axisMetrics: HashMap<String, Any>, s: ChartSeries): String {
        return ""
    }

    /**
     * returns the SVG-ready data label for the given set of data label options
     * TODO: get separator character
     *
     * @param datalabel  int[] Data label options (indexed by current series # s)
     * <br></br>can be one or more of:
     * <br></br>VALUELABEL= 0x1;
     * <br></br>VALUEPERCENT= 0x2;
     * <br></br>CATEGORYPERCENT= 0x4;
     * <br></br>SMOOTHEDLINE= 0x8;
     * <br></br>CATEGORYLABEL= 0x10;
     * <br></br>BUBBLELABEL= 0x20;
     * <br></br>SERIESLABEL= 0x40;
     * @param series     ArrayList of series and category data: structure:  category data, series 0-n data, series colors, series Labels (Eventually will be an object)
     * @param val        double current series value
     * @param percentage double percentage value (pie charts only)
     * @param s          int current series #
     * @param cat        string current cat
     * @return
     */
    fun getSVGDataLabels(datalabels: IntArray, axisMetrics: HashMap<String, Any>, `val`: Double, percentage: Double, s: Int, legends: Array<String>, cat: String): String? {

        if (s >= datalabels.size)
        // can happen with Pie-style charts
            return null
        val showValueLabel = datalabels[s] and AttachedLabel.VALUELABEL == AttachedLabel.VALUELABEL
        val showValuePercent = datalabels[s] and AttachedLabel.VALUEPERCENT == AttachedLabel.VALUEPERCENT
        val showCatPercent = datalabels[s] and AttachedLabel.CATEGORYPERCENT == AttachedLabel.CATEGORYPERCENT
        val showCategories = datalabels[s] and AttachedLabel.CATEGORYLABEL == AttachedLabel.CATEGORYLABEL
        val showBubbleLabel = datalabels[s] and AttachedLabel.BUBBLELABEL == AttachedLabel.BUBBLELABEL
        val showValue = datalabels[s] and AttachedLabel.VALUE == AttachedLabel.VALUE
        if (showValue || showCategories || showValueLabel || showValuePercent || showBubbleLabel) {
            var l = ""
            if (showValueLabel)
                l += legends[s] + " "    // series names
            if (showCategories)
                l += CellFormatFactory.fromPatternString(
                        axisMetrics["xPattern"] as String).format(cat) + " "    // categories
            if (showValue || showBubbleLabel) {
                l += CellFormatFactory.fromPatternString(
                        axisMetrics["yPattern"] as String).format(`val`.toString())
                /*try {
					int v= new Double(val).intValue();
					if (v==val)
						l+=v + " ";
					else
						l+=val + " ";
				} catch (Exception e) {
					l+= val + " ";
				}*/
            }
            if (showValuePercent)
                l += Math.round(percentage * 100).toInt().toString() + "%"
            return l
        }
        return null
    }

    /** Show or remove Data Table for Chart
     * NOTE:  METHOD IS STILL EXPERIMENTAL
     * @param bShow
     *
     * public void showDataTable(boolean bShow) {
     * int i= Chart.findRecPosition(chartArr, Dat.class);
     * if (bShow) {
     * if (i==-1) { // add Dat
     * Dat d= (Dat)Dat.getPrototype(true); // create data table
     * i= Chart.findRecPosition(chartArr, AxisParent.class);
     * this.chartArr.add(++i, d);
     * }
     * } else if (i > 0) {
     * chartArr.remove(i);	// remove Dat - Data Table options + all associated recs
     * }
     * }
     */


    /**
     * return truth if Chart has a data legend key showing
     *
     * @return
     */
    fun hasDataLegend(): Boolean {
        return dataLegend != null
    }

    /**
     * show or hide chart legend key
     *
     * @param bShow    boolean show or hide
     * @param vertical boolean show as vertical or horizontal
     */
    fun showLegend(bShow: Boolean, vertical: Boolean) {
        if (bShow && dataLegend == null) {
            dataLegend = Legend.createDefaultLegend(wb)
            dataLegend!!.parentChart = chartobj.parentChart
            for (j in dataLegend!!.chartArr.indices)
                (dataLegend!!.chartArr[j] as GenericChartObject).parentChart = chartobj.parentChart
            dataLegend!!.setVertical(vertical)
            cf!!.chartArr.add(dataLegend)
        } else if (bShow) {
            dataLegend!!.setVertical(vertical)
        } else if (dataLegend != null) {
            val i = Chart.findRecPosition(cf!!.chartArr, Legend::class.java)
            cf!!.chartArr.removeAt(i)
            dataLegend = null
        }
    }

    /**
     * sets this chart (Area, Line, Scatter) to have drop lines
     */
    fun setHasDropLines() {
        cf!!.setHasDropLines()
    }

    /**
     * look up a generic chart option
     *
     * @param op
     * @return
     */
    fun getChartOption(op: String): String? {
        return chartobj.getChartOption(op) ?: return cf!!.getChartOption(op)
    }

    /**
     * interface for setting chart-type-specific options
     * in a generic fashion
     *
     * @param op  option
     * @param val value
     * @see OpenXLS.handleChartElement
     *
     * @see ChartHandle.getXML
     */
    fun setChartOption(op: String, `val`: String): Boolean {
        if (!chartobj.setChartOption(op, `val`))
        // if not handled,
            cf!!.setOption(op, `val`)
        return true
    }

    /**
     * returns the 3d record of the desired chart
     *
     * @param chartType one of the chart type constants
     * @return
     */
    fun initThreeD(charttype: Int): ThreeD {
        /* TODO: test if this is necessary        AxisParent ap = this.getAxisParent();
    	try {
        	// first thing, remove the PlotArea and Frame - don't know why but is necessary!!
    		int x= Chart.findRecPosition(ap.chartArr, PlotArea.class);
    		if (x!=-1) ap.chartArr.remove(x);
    		x= Chart.findRecPosition(ap.chartArr, Frame.class);
    		if (x!=-1) ap.chartArr.remove(x);
    	} catch (Exception e) {}*/
        val td = cf!!.getThreeDRec(true)
        td!!.setIsPie(charttype == ChartConstants.PIECHART || charttype == ChartConstants.DOUGHNUTCHART)
        return td
    }

    /**
     * return the 3d record of the chart
     * <br></br>Creates if not present, if bCreate is true
     *
     * @return
     */
    fun getThreeDRec(bCreate: Boolean): ThreeD? {
        return cf!!.getThreeDRec(bCreate)
    }

    fun setIs100Psercent(ispercentage: Boolean) {
        chartobj.is100Percent = ispercentage
    }

    fun addLegend() {
        if (dataLegend == null) {
            // TODO: ADD LEGEND to cf
            //int i = Chart.findRecPosition(cf.chartArr, Legend.class);
        }
    }


    /**
     * returns the SVG for the font style of this object
     *
     * @param font size in pt
     */
    protected fun getFontSVG(sz: Int): String {
        //		io.starter.formats.XLS.Font f = this.getFont();	// this is not correct!
        //	if(f == null) // return a default
        return "font-family='Arial' font-size='$sz' fill='$darkColor' "

        //return f.getSVG();
    }

    /**
     * Convert the user-friendly OOXML shape string to 2003-v int shape flag
     *
     * @param shape
     * @return
     */
    fun convertShape(shape: String): Int {
        defaultShape = 0
        if (shape == "box")
            defaultShape = 0
        if (shape == "cone")
        // 1 1
            defaultShape = 257
        if (shape == "coneToMax")
        // 1 2
            defaultShape = 513
        if (shape == "cylinder")
        // 1 0
            defaultShape = 1
        if (shape == "pyramid")
        // 0 1
            defaultShape = 256
        if (shape == "pyramidToMax")
        // 0 2
            defaultShape = 512
        return defaultShape
    }

    /**
     * @param fName
     */
    fun WriteMainChartRecs(fName: String) {
        class util {
            @Throws(IOException::class)
            fun writeRecs(b: BiffRec?, writer: BufferedWriter, level: Int) {
                val tabs = "\t\t\t\t\t\t\t\t\t\t"
                if (b == null) return
                writer.write(tabs.substring(0, level) + b.javaClass.toString().substring(b.javaClass.toString().lastIndexOf('.') + 1))
                if (b is io.starter.formats.XLS.charts.SeriesText)
                    writer.write("\t[$b]")
                else if (b is MSODrawing) {
                    writer.write("\t[$b]")
                    //								writer.write("\t[" + ByteTools.getByteDump(b.getData(), 0).substring(11)+ "]");
                    writer.write(b.debugOutput())
                    writer.write("\t[" + ByteTools.getByteDump(b.data, 0).substring(11) + "]")
                } else if (b is Obj) {
                    writer.write(b.debugOutput())
                } else if (b is Label) {
                    writer.write("\t[" + b.stringVal + "]")
                } else
                // all else, write bytes
                    writer.write("\t[" + ByteTools.getByteDump(ByteTools.shortToLEBytes(b.opcode), 0) + "][" + ByteTools.getByteDump(b.data, 0).substring(11) + "]")
                writer.newLine()
                try {
                    if ((b as GenericChartObject).chartArr.size > 0) {
                        val chartArr = b.chartArr
                        for (i in chartArr.indices) {
                            writeRecs(chartArr[i], writer, level + 1)

                        }
                    }
                } catch (ce: ClassCastException) {
                }

            }
        }


        try {
            val f = java.io.File(fName)
            var writer: BufferedWriter? = BufferedWriter(FileWriter(f))
            val u = util()

            val v = this.parentChart!!.allSeries
            for (i in v.indices) {
                u.writeRecs(v[i] as BiffRec, writer, 0)
            }
            writer!!.newLine()
            val chartArr = this.cf!!.chartArr
            for (i in chartArr.indices) {
                u.writeRecs(chartArr[i], writer, 0)
            }

            writer.flush()
            writer.close()
            writer = null
        } catch (e: Exception) {
        }

    }

    companion object {
        /**
         * serialVersionUID
         */
        private const val serialVersionUID = -7862828186455339066L

        /**
         * creates a new chart type object of the desired chart type.  Will set options if already set via ChartFormat
         *
         * @param chartType
         * @param parentChart
         * @param cf
         * @return
         */
        fun create(chartType: Int, parentChart: Chart, cf: ChartFormat): ChartType? {
            val co = ChartType.createUnderlyingChartObject(chartType, parentChart, cf)    //, ((ChartType)chartobj.get(0)));

            return createChartTypeObject(co, cf, parentChart.workBook)
        }


        /**
         * create and return the appropriate ChartType for the given
         * chart record (Bar, Radar, Line, etc)
         *
         * @param ch GenericChartObject representing the Chart Type
         * <br></br>Must be one of:  Area, Bar, Line, Surface, Pie, Radar, Scatter, BopPop, RadarArea
         * @return ChartType chart object
         */
        fun createChartTypeObject(ch: GenericChartObject?, cf: ChartFormat?, wb: WorkBook?): ChartType? {
            if (cf == null)
            // TODO: throw exception
                return null

            val barshape = cf.barShape
            val threeD = cf.isThreeD(ch!!.chartType)
            val isStacked = ch.isStacked
            val is100Percent = ch.is100Percent

            when (ch.chartType) {
                ChartConstants.COLCHART -> {
                    if (barshape == ChartConstants.SHAPEDEFAULT) { // regular column
                        if (isStacked || is100Percent)
                            return StackedColumn(ch, cf, wb)
                        return if (threeD) Col3DChart(ch, cf, wb) else ColChart(ch, cf, wb)
                    } else if (barshape == ChartConstants.SHAPECONE) { // Cone chart	 always 3d
                        return ConeChart(ch, cf, wb)
                    } else if (barshape == ChartConstants.SHAPECYLINDER) { // Cylinder chart always 3d
                        return CylinderChart(ch, cf, wb)
                    } else if (barshape == ChartConstants.SHAPEPYRAMID) {    // Pyramid chart	alwasy 3d
                        return PyramidChart(ch, cf, wb)
                    }
                    if (barshape == ChartConstants.SHAPEDEFAULT) { // regular Bar
                        return if (threeD) Bar3DChart(ch, cf, wb) else BarChart(ch, cf, wb)
                    } else if (barshape == ChartConstants.SHAPECONE) { // ConeBarchart	always 3d
                        return ConeBarChart(ch, cf, wb)
                    } else if (barshape == ChartConstants.SHAPECYLINDER) { // CylinderBar chart	always 3d
                        return CylinderBarChart(ch, cf, wb)
                    } else if (barshape == ChartConstants.SHAPEPYRAMID) {    // PyramidBar chart	always 3d
                        return PyramidBarChart(ch, cf, wb)
                    }
                    return if (threeD) Line3DChart(ch, cf, wb) else LineChart(ch, cf, wb)
                }
                ChartConstants.BARCHART -> {
                    if (barshape == ChartConstants.SHAPEDEFAULT) {
                        return if (threeD) Bar3DChart(ch, cf, wb) else BarChart(ch, cf, wb)
                    } else if (barshape == ChartConstants.SHAPECONE) {
                        return ConeBarChart(ch, cf, wb)
                    } else if (barshape == ChartConstants.SHAPECYLINDER) {
                        return CylinderBarChart(ch, cf, wb)
                    } else if (barshape == ChartConstants.SHAPEPYRAMID) {
                        return PyramidBarChart(ch, cf, wb)
                    }
                    return if (threeD) Line3DChart(ch, cf, wb) else LineChart(ch, cf, wb)
                }
                ChartConstants.LINECHART -> {
                    return if (threeD) Line3DChart(ch, cf, wb) else LineChart(ch, cf, wb)
                }
                ChartConstants.STOCKCHART -> return StockChart(ch, cf, wb)
                ChartConstants.PIECHART -> {
                    cf.setPercentage(0)
                    return if (threeD) Pie3dChart(ch, cf, wb) else PieChart(ch, cf, wb)
                }
                ChartConstants.AREACHART -> {
                    if (ch.isStacked)
                        return StackedAreaChart(ch, cf, wb)
                    return if (threeD) Area3DChart(ch, cf, wb) else AreaChart(ch, cf, wb)
                }
                ChartConstants.SCATTERCHART -> return ScatterChart(ch, cf, wb)
                ChartConstants.RADARCHART -> return RadarChart(ch, cf, wb)
                ChartConstants.SURFACECHART -> {
                    return if ((ch as Surface).is3d) Surface3DChart(ch, cf, wb) else SurfaceChart(ch, cf, wb)
                }
                ChartConstants.DOUGHNUTCHART -> {
                    cf.setPercentage(0)
                    return DoughnutChart(ch, cf, wb)
                }
                ChartConstants.BUBBLECHART -> return BubbleChart(ch, cf, wb)
                ChartConstants.RADARAREACHART -> return RadarAreaChart(ch, cf, wb)
                ChartConstants.OFPIECHART -> {
                    cf.setPercentage(0)
                    return OfPieChart(ch, cf, wb)
                }
                else -> return null
            }
        }

        /**
         * create a new low-level chart object which determines the ChartTypeObject, and update the existing chartobj
         *
         * @param chartType
         * @param parentchart
         * @param chartgroup
         * @return
         */
        fun createUnderlyingChartObject(chartType: Int, parentchart: Chart, cf: ChartFormat?): GenericChartObject? {//, ChartType chartobj) {
            var c: GenericChartObject? = null
            when (chartType) {
                ChartConstants.COLCHART // column-type
                -> {
                    val col = Bar.prototype as Bar?
                    col!!.setAsColumnChart()
                    c = col
                }

                ChartConstants.BARCHART // bar-type
                -> {
                    val bar = Bar.prototype as Bar?
                    bar!!.setAsBarChart()
                    c = bar
                }

                ChartConstants.PIECHART // Pie
                -> {
                    val p = Pie.prototype as Pie?
                    p!!.setAsPieChart()
                    c = p
                }

                ChartConstants.STOCKCHART -> {
                    val st = Line.prototype as Line?
                    st!!.setAsStockChart()
                    c = st
                }

                ChartConstants.LINECHART -> {
                    val l = Line.prototype as Line?
                    c = l
                }

                ChartConstants.AREACHART -> {
                    val a = Area.prototype as Area?
                    c = a
                }

                ChartConstants.SCATTERCHART -> {
                    val s = Scatter.prototype as Scatter?
                    s!!.setAsScatterChart()
                    c = s
                }

                ChartConstants.RADARCHART -> {
                    val r = Radar.prototype as Radar?
                    c = r
                }

                ChartConstants.SURFACECHART -> {
                    val su = Surface.prototype as Surface?
                    c = su
                }

                ChartConstants.DOUGHNUTCHART -> {
                    val d = Pie.prototype as Pie?
                    d!!.setAsDoughnutChart()
                    c = d
                }

                ChartConstants.BUBBLECHART -> {
                    val bu = Scatter.prototype as Scatter?
                    bu!!.setAsBubbleChart()
                    c = bu
                }

                ChartConstants.RADARAREACHART -> {
                    val ra = RadarArea.prototype as RadarArea?
                    c = ra
                }

                // note that, for the below chart types, the underlying type will be either COL or BAR
                // which actual chart is determined also by the bar shape
                ChartConstants.PYRAMIDCHART -> {
                    val pyramid = Bar.prototype as Bar?
                    pyramid!!.setAsColumnChart()
                    cf!!.barShape = ChartConstants.SHAPEPYRAMID
                    c = pyramid
                }

                ChartConstants.CONECHART -> {
                    val cone = Bar.prototype as Bar?
                    cone!!.setAsColumnChart()
                    cf!!.barShape = ChartConstants.SHAPECONE
                    c = cone
                }

                ChartConstants.CYLINDERCHART -> {
                    val cy = Bar.prototype as Bar?
                    cy!!.setAsColumnChart()
                    cf!!.barShape = ChartConstants.SHAPECYLINDER
                    c = cy
                }

                ChartConstants.PYRAMIDBARCHART -> {
                    val pb = Bar.prototype as Bar?
                    pb!!.setAsBarChart()
                    cf!!.barShape = ChartConstants.SHAPEPYRAMID
                    c = pb
                }

                ChartConstants.CONEBARCHART -> {
                    val cb = Bar.prototype as Bar?
                    cb!!.setAsBarChart()
                    cf!!.barShape = ChartConstants.SHAPECONE
                    c = cb
                }

                ChartConstants.CYLINDERBARCHART -> {
                    val cyb = Bar.prototype as Bar?
                    cyb!!.setAsBarChart()
                    cf!!.barShape = ChartConstants.SHAPECYLINDER
                    c = cyb
                }

                ChartConstants.OFPIECHART -> {
                    val ofpie = Boppop.prototype as Boppop?
                    c = ofpie
                }
            }
            if (c != null) {
                c.parentChart = parentchart
            }
            return c
        }


        /**
         * parse the chart object OOXML element (barchart, area3DChart, etc.) and create the corresponding chart type objects
         *
         * @param xpp         XmlPullParser
         * @param wbh         workBookHandle
         * @param parentChart parent Chart Object
         * @param nChart      chart grouping number, 0 for default, 1-9 for overlay
         * @return
         */
        fun parseOOXML(xpp: XmlPullParser, wbh: WorkBookHandle, parentChart: Chart, nChart: Int): ChartType? {
            try {
                val endTag = xpp.name
                var tnm = xpp.name
                var eventType = xpp.eventType
                var chartType = ChartConstants.BARCHART
                val ca = parentChart.axes

                val lastTag = java.util.Stack()        // keep track of element hierarchy
                val cf = parentChart.getChartOjectParent(nChart)

                val options = EnumSet.noneOf<ChartOptions>(ChartOptions::class.java)    // chart-specific options such as threed, stacked ...
                if (tnm == "bubble3D") {    // bubble3D tag appears for each series in 3D bubble chart
                    options.add(ChartOptions.THREED)
                } else if (tnm == OOXMLConstants.twoDchartTypes[ChartHandle.BARCHART]) {
                    chartType = ChartConstants.BARCHART
                } else if (tnm == OOXMLConstants.twoDchartTypes[ChartHandle.LINECHART]) {
                    chartType = ChartConstants.LINECHART
                } else if (tnm == OOXMLConstants.twoDchartTypes[ChartHandle.PIECHART]) {
                    chartType = ChartConstants.PIECHART
                    parentChart.axes!!.removeAxes()
                } else if (tnm == OOXMLConstants.twoDchartTypes[ChartHandle.AREACHART]) {
                    chartType = ChartConstants.AREACHART
                } else if (tnm == OOXMLConstants.twoDchartTypes[ChartHandle.SCATTERCHART]) {
                    chartType = ChartConstants.SCATTERCHART
                    ca!!.removeAxis(ChartConstants.XAXIS)
                    ca.createAxis(ChartConstants.XVALAXIS)
                } else if (tnm == OOXMLConstants.twoDchartTypes[ChartHandle.RADARCHART]) {
                    chartType = ChartConstants.RADARCHART
                } else if (tnm == OOXMLConstants.twoDchartTypes[ChartHandle.SURFACECHART]) {
                    chartType = ChartConstants.SURFACECHART
                } else if (tnm == OOXMLConstants.twoDchartTypes[ChartHandle.DOUGHNUTCHART]) {
                    chartType = ChartConstants.DOUGHNUTCHART
                    parentChart.axes!!.removeAxes()
                } else if (tnm == OOXMLConstants.twoDchartTypes[ChartHandle.BUBBLECHART]) {
                    chartType = ChartConstants.BUBBLECHART
                    ca!!.removeAxis(ChartConstants.XAXIS)
                    ca.createAxis(ChartConstants.XVALAXIS)
                } else if (tnm == OOXMLConstants.threeDchartTypes[ChartHandle.BARCHART]) {
                    chartType = ChartConstants.BARCHART
                    options.add(ChartOptions.THREED)
                } else if (tnm == OOXMLConstants.threeDchartTypes[ChartHandle.LINECHART]) {
                    chartType = ChartConstants.LINECHART
                    options.add(ChartOptions.THREED)
                } else if (tnm == OOXMLConstants.threeDchartTypes[ChartHandle.PIECHART]) {
                    chartType = ChartConstants.PIECHART
                    options.add(ChartOptions.THREED)
                    parentChart.axes!!.removeAxes()
                } else if (tnm == OOXMLConstants.threeDchartTypes[ChartHandle.AREACHART]) {
                    chartType = ChartConstants.AREACHART
                    options.add(ChartOptions.THREED)
                } else if (tnm == OOXMLConstants.threeDchartTypes[ChartHandle.SCATTERCHART]) {
                    chartType = ChartConstants.SCATTERCHART
                    ca!!.removeAxis(ChartConstants.XAXIS)
                    ca.createAxis(ChartConstants.XVALAXIS)
                    options.add(ChartOptions.THREED)
                } else if (tnm == OOXMLConstants.threeDchartTypes[ChartHandle.RADARCHART]) {
                    chartType = ChartConstants.RADARCHART
                    options.add(ChartOptions.THREED)
                } else if (tnm == OOXMLConstants.threeDchartTypes[ChartHandle.SURFACECHART]) {
                    chartType = ChartConstants.SURFACECHART
                    options.add(ChartOptions.THREED)
                } else if (tnm == OOXMLConstants.threeDchartTypes[ChartHandle.DOUGHNUTCHART]) {
                    chartType = ChartConstants.DOUGHNUTCHART
                    options.add(ChartOptions.THREED)
                    parentChart.axes!!.removeAxes()
                } else if (tnm == OOXMLConstants.twoDchartTypes[ChartHandle.OFPIECHART]) {
                    chartType = ChartConstants.OFPIECHART
                    ca!!.removeAxis(ChartConstants.XAXIS)
                    ca.removeAxis(ChartConstants.YAXIS)
                } else if (tnm == OOXMLConstants.twoDchartTypes[ChartHandle.STOCKCHART]) {
                    chartType = ChartConstants.STOCKCHART
                }

                val co = ChartType.createUnderlyingChartObject(chartType, parentChart, cf)    //, ((ChartType)chartobj.get(0)));
                cf!!.setChartObject(co)    // sets the chart format (parent of chart item) to the specific chart item
                // exception for surface charts in 3d
                if (chartType == ChartConstants.SURFACECHART && tnm == OOXMLConstants.threeDchartTypes[ChartHandle.SURFACECHART])
                    (co as Surface).is3d = true
                var ct = ChartType.createChartTypeObject(co, cf, parentChart.workBook)
                parentChart.addChartType(ct, nChart)

                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        tnm = xpp.name
                        lastTag.push(tnm)
                        var v: String? = null
                        try {
                            v = xpp.getAttributeValue(0)
                        } catch (e: IndexOutOfBoundsException) {
                        }

                        // TODO:
                        // hasMarkers ****
                        // dLbls
                        // smooth	-- line
                        // marker -- 	line
                        // bandfmts -- surface, surface3d
                        // bubble3D,
                        // bubbleScale
                        // showNegBubbles
                        // sizeRepresents
                        // custSplit	-- OfPie
                        if (tnm == "grouping") {    //This element specifies the type of grouping for a column, line, or area chart + 3d versions
                            if (v == "stacked")
                                co!!.isStacked = true
                            else if (v == "percentStacked")
                                co!!.is100Percent = true
                            else if (v == "clustered")
                                cf.setIsClustered(true)        // bar/col only
                            else if (v == "standard") {    // for Line, Line3d, Area, Area3d and Bar3d, Col3d - does not appear valid for Bar 2d and Col 2d
                                co!!.isStacked = false
                                co.is100Percent = false
                            }
                        } else if (tnm == "barDir") {    //
                            if (v == "col") {
                                (co as Bar).setAsColumnChart()
                                ct = ChartType.createChartTypeObject(co, cf, parentChart.workBook)
                                parentChart.addChartType(ct, nChart)
                            }
                        } else if (tnm == "shape") {    // bar3d
                            cf.barShape = ct!!.convertShape(v!!)
                            if (ct.defaultShape != 0) {
                                ct = ChartType.createChartTypeObject(co, cf, parentChart.workBook)
                                parentChart.addChartType(ct, nChart)
                            }
                        } else if (tnm == "radarStyle") {
                            if (v == "filled") {
                                (ct as RadarChart).isFilled = true
                            } else if (v == "marker") {

                            }
                        } else if (tnm == "wireframe") {    // surface
                            (co as Surface).isWireframe = v != null && v == "1"
                        } else if (tnm == "scatterStyle") {
                            if (v == "lineMarker") {
                                cf.setHasLines()
                            } else if (v == "smoothMarker") {
                                cf.hasSmoothLines = true
                            } else if (v == "marker")
                            else if (v == "line")
                                cf.setHasLines()
                            else if (v == "smooth")
                                cf.hasSmoothLines = true//cf.seth;
                        } else if (tnm == "varyColors") {
                            cf.varyColor = xpp.getAttributeValue(0) == "1"
                        } else if (tnm == "dropLines") {
                            val cl = cf.addChartLines(ChartLine.TYPE_DROPLINE.toInt())
                            cl.parseOOXML(xpp, lastTag, cf, wbh)
                        } else if (tnm == "hiLowLines") {
                            val cl = cf.addChartLines(ChartLine.TYPE_HILOWLINE.toInt())
                            cl.parseOOXML(xpp, lastTag, cf, wbh)
                        } else if (tnm == "upDownBars") {
                            cf.parseUpDownBarsOOXML(xpp, lastTag, wbh)
                        } else if (tnm == "serLines") {
                            val cl = cf.addChartLines(ChartLine.TYPE_SERIESLINE.toInt())
                            cl.parseOOXML(xpp, lastTag, cf, wbh)
                        } else if (tnm == "overlap") {    // bar
                            co!!.setChartOption("Overlap", v)
                        } else if (tnm == "gapWidth") {
                            co!!.setChartOption("Gap", v)        // bar
                        } else if (tnm == "ofPieType") {
                            (co as Boppop).isPieOfPie = "pie" == v
                        } else if (tnm == "gapDepth") {
                            cf.gapDepth = Integer.valueOf(v!!)    // bar3d, area3d, line3d
                        } else if (tnm == "firstSliceAn") {
                            (co as Pie).setAnStart(Integer.valueOf(v!!))        // pie or doughnut
                        } else if (tnm == "holeSize") {
                            (co as Pie).doughnutSize = Integer.valueOf(v!!)        // pie or doughnut
                        } else if (tnm == "secondPieSize") {
                            (co as Boppop).secondPieSize = Integer.valueOf(v!!)    // OfPie (Pie of Pie, Bar of Pie)
                        } else if (tnm == "splitType") {
                            (co as Boppop).setSplitType(v)    // OfPie (Pie of Pie, Bar of Pie)
                        } else if (tnm == "splitPos") {
                            (co as Boppop).splitPos = Integer.valueOf(v!!)    // OfPie (Pie of Pie, Bar of Pie)
                        } else if (tnm == "ser") {
                            val s = ChartSeries.parseOOXML(xpp, wbh, ct, false, lastTag)
                        } else if (tnm == "dLbls") {
                        } else if (tnm == "marker") {        // line only?
                            //				       	 	m= (Marker) Marker.parseOOXML(xpp, lastTag).cloneElement();
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        if (xpp.name == endTag)
                            break

                    }
                    eventType = xpp.next()
                }
                return ct

                //		    lastTag.pop();	// chart type tag will be added in parseOOXML
                //mychart.getChartObject(nChart).parseOOXML(xpp, this.wbh, lastTag);
            } catch (e: Exception) {
                Logger.logErr("ChartType.parseChartType: $e")
            }

            return null
        }

        // SVG Convenience Methods
        fun getScript(range: String): String {
            return "onmouseover='highLight(evt); showRange(\"$range\");' onclick='handleClick(evt);' onmouseout='restore(evt); hideRange();'"
            //		return "onmouseover='highLight(evt);' onclick='handleClick(evt);' onmouseout='restore(evt);'";
        }

        val fillOpacity: String
            get() = ".75"


        val textColor: String
            get() = "#222222"

        val lightColor: String
            get() = "#CCCCCC"

        val mediumColor: String
            get() = "#555555"


        val darkColor: String
            get() = "#333333"

        /**
         * Returns SVG used to define font for data labels
         * TODO: read correct value from chart recs
         *
         * @return
         */
        val dataLabelFontSVG: String
            get() {
                val sz = 9
                return "font-family='Arial' font-size='" + sz + "' fill='" + ChartType.darkColor + "' "
            }

        /**
         * returns the SVG for the font style of this object
         *
         * @param stroke size in pt
         */
        val strokeSVG: String
            get() = getStrokeSVG(1f, mediumColor)

        /**
         * returns the SVG for the font style of this object
         *
         * @param stroke size in pt
         * @param String stroke color in HTML format
         */
        fun getStrokeSVG(sz: Float, strokeclr: String): String {
            return " stroke='$strokeclr'  stroke-opacity='1' stroke-width='$sz' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'"
        }
    }
}
