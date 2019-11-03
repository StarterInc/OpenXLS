/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or
 * modify
 * it under the terms of the GNU Lesser General Public
 * License as
 * published by the Free Software Foundation, either version
 * 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be
 * useful,
 * but WITHOUT ANY WARRANTY; without even the implied
 * warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
 * the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General
 * Public
 * License along with OpenXLS. If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package io.starter.OpenXLS

import io.starter.formats.OOXML.*
import io.starter.formats.XLS.*
import io.starter.formats.XLS.charts.*
import io.starter.toolkit.Logger
import org.json.JSONArray
import org.json.JSONException
import org.json.JSONObject
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserFactory

import java.util.ArrayList
import java.util.EnumSet
import java.util.HashMap
import java.util.Vector

/**
 * Chart Handle allows for manipulation of Charts within a WorkBook. <br></br>
 * Allows for run-time modification of Chart titles, labels, categories, and
 * data Cells. <br></br>
 * Modification of Chart data cells allows you to completely modify the data
 * shown on the Chart. <br></br>
 * <br></br>
 * To create a new chart and add to a worksheet, use:
 * ChartHandle.createNewChart(sheet, charttype, chartoptions). <br></br>
 * <br></br>
 * To Obtain an array of existing chart handles, use <br></br>
 * WorkBookHandle.getCharts() <br></br>
 * or WorkBookHandle.getCharts(chart title)
 *
 * [Starter Inc.](http://starter.io)
 *
 * @see WorkBookHandle
 */
class ChartHandle
/**
 * Constructor which creates a new ChartHandle from an existing Chart Object
 *
 * @param Chart          c - the source Chart object
 * @param WorkBookHandle wb - the parent WorkBookHandle
 */
(private var mychart: Chart?, wb: WorkBookHandle) : ChartConstants {

    var workBookHandle: WorkBookHandle
        protected set

    /**
     * Returns the title of the Chart
     *
     * @return String title of the Chart
     */
    /**
     * Sets the title of the Chart
     *
     * @param String title - Chart title
     */
    var title: String?
        get() = mychart!!.title
        set(title) {
            this.mychart!!.title = title
        }

    /**
     * returns the data range used by the chart
     *
     * @return
     */
    val dataRangeJSON: String
        get() = this.mychart!!.chartSeries.dataRangeJSON.toString()

    val encompassingDataRange: IntArray?
        get() = getEncompassingDataRange(this.mychart!!.chartSeries
                .dataRangeJSON)

    /**
     * returns the ordinal id associated with the underlying Chart Object
     *
     * @return int chart id
     * @see WorkBookHandle.getChartById
     */
    val id: Int
        get() = this.mychart!!.id

    /** */
    /**
     * Returns an ordered array of strings representing all the series ranges in the
     * Chart. <br></br>
     * Each series can only represent one bar, line or wedge of data.
     *
     * @return String[] each item being a Cell Range representing one bar, line or
     * wedge of data in the Chart
     * @see ChartHandle.getCategories
     */
    // -1 is flag for all rather than for a
    // specific chart
    val series: Array<String>
        get() = mychart!!.getSeries(-1)

    /**
     * Returns an ordered array of strings representing all the category ranges in
     * the chart. <br></br>
     * This vector corresponds to the getSeries() method so will often contain
     * duplicates, as while the series data changes frequently, category data is the
     * same throughout the chart.
     *
     * @return String[] each item being a Cell Range representing the Category Data
     * @see ChartHandle.getSeries
     */
    // -1 is flag for all rather than for a
    // specific chart
    val categories: Array<String>
        get() = getCategories(-1)

    /**
     * Returns an array of ChartSeriesHandle Objects, one for each bar, line or
     * wedge of data.
     *
     * @return ChartSeriesHandle[] Array of ChartSeriesHandle Objects representing
     * Chart Series Data (Series and Categories)
     * @see ChartSeriesHandle
     */
    // get ALL
    val allChartSeriesHandles: Array<ChartSeriesHandle>
        get() = getAllChartSeriesHandles(-1)

    /**
     * returns the Y axis Title
     *
     * @return String title
     */
    /**
     * Sets the Y axis Title
     *
     * @param String yTitle - new Y Axis title
     */
    var yAxisLabel: String
        get() = mychart!!.axes!!.getTitle(YAXIS)
        set(yTitle) = mychart!!.axes!!.setTitle(YAXIS, yTitle)

    /**
     * returns the X axis Title
     *
     * @return String title
     */
    /**
     * Sets the XAxisTitle
     *
     * @param String xTitle - new X Axis title
     */
    var xAxisLabel: String
        get() = mychart!!.axes!!.getTitle(XAXIS)
        set(xTitle) = mychart!!.axes!!.setTitle(XAXIS, xTitle)

    /**
     * returns the Z axis Title, if any
     *
     * @return String Title
     */
    /**
     * set the Z AxisTitle
     *
     * @param String zTitle - new Z Axis Title
     */
    var zAxisLabel: String
        get() = mychart!!.axes!!.getTitle(ZAXIS)
        set(zTitle) = mychart!!.axes!!.setTitle(ZAXIS, zTitle)

    /**
     * Returns true if the Y Axis (Value axis) is set to automatic scale
     *
     *
     * The default setting for charts is known as Automatic Scaling <br></br>
     * When data changes, the chart automatically adjusts the scale (minimum,
     * maximum values plus major and minor tick units) as necessary
     *
     * @return boolean true if Automatic Scaling is turned on
     * @see getAxisAutomaticScale
     */
    /**
     * Sets the automatic scale option on or off for the Y Axis (Value axis)
     *
     *
     * Automatic Scaling will automatically set the scale maximum, minimum and tick
     * units upon data changes, and is the default chart setting
     *
     * @param b
     * @see setAxisAutomaticScale
     */
    var axisAutomaticScale: Boolean
        get() = mychart!!.axisAutomaticScale
        set(b) {
            mychart!!.axes!!.axisAutomaticScale = b
            mychart!!.setDirtyFlag(true)
        }

    /**
     * Returns the minimum value of the Y Axis (Value Axis) scale
     *
     * @return int Miminum Scale value for Y axis
     * @see getAxisMinScale
     */
    // -1 is flag for all
    // rather than for a
    // specific chart
    val axisMinScale: Double
        get() {
            val minmax = mychart!!.getMinMax(this.workBookHandle)
            return mychart!!.axes!!.getMinMax(minmax[0], minmax[1])[0]
        }

    /**
     * Returns the maximum value of the Y Axis (Value Axis) scale
     *
     * @return int Maximum Scale value for Y axis
     * @see getAxisMaxScale
     */
    val axisMaxScale: Double
        get() {
            val minmax = mychart!!.getMinMax(this.workBookHandle)
            return mychart!!.axes!!.getMinMax(minmax[0], minmax[1])[1]
        }

    /**
     * Returns the major tick unit of the Y Axis (Value Axis)
     *
     * @return int major tick unit
     * @see getAxisMajorUnit
     */
    val axisMajorUnit: Int
        get() {
            val minmax = mychart!!.getMinMax(this.workBookHandle)
            return mychart!!.axes!!.getMinMax(minmax[0], minmax[1])[2].toInt()
        }

    /**
     * Returns the minor tick unit of the Y Axis (Value Axis)
     *
     * @return int minor tick unit
     * @see getAxisMinorUnit
     */
    val axisMinorUnit: Int
        get() {
            val minmax = mychart!!.getMinMax(this.workBookHandle)
            return mychart!!.axes!!.getMinMax(minmax[0], minmax[1])[1].toInt()
        }

    /**
     * returns the Font associated with the Chart Title
     *
     * @return io.starter.Formats.XLS.Font
     * @see Font
     */
    /**
     * set the font for the Chart Title
     *
     * @param io.starter.Formats.XLS.Font f - desired font for the Chart Title
     * @see Font
     */
    // font doesn't exist yet, add to streamer
    // flag to insert
    var titleFont: Font?
        get() = mychart!!.titleFont
        set(f) {
            var idx = workBookHandle.workBook!!.getFontIdx(f)
            if (idx == -1) {
                f.idx = -1
                idx = workBookHandle.workBook!!.insertFont(f) + 1
            }
            mychart!!.setTitleFont(idx)
        }

    /**
     * return the Font associated with the Chart Axes
     *
     * @return io.starter.Formats.XLS.Font for Chart Axes or null if no axes
     * @see Font
     */
    /**
     * set the font for all axes on the Chart
     *
     * @param io.starter.Formats.XLS.Font f - desired font for the Chart axes
     * @see Font
     */
    // font doesn't exist yet, add to streamer
    // flag to insert
    var axisFont: Font?
        get() {
            var f: Font? = null
            f = mychart!!.axes!!.getTitleFont(XAXIS)
            if (f != null)
                return f
            f = mychart!!.axes!!.getTitleFont(YAXIS)
            if (f != null)
                return f
            f = mychart!!.axes!!.getTitleFont(ZAXIS)
            return f
        }
        set(f) {
            var idx = workBookHandle.workBook!!.getFontIdx(f)
            if (idx == -1) {
                f.idx = -1
                idx = workBookHandle.workBook!!.insertFont(f) + 1
            }
            mychart!!.axes!!.setTitleFont(XAXIS, idx)
            mychart!!.axes!!.setTitleFont(YAXIS, idx)
            mychart!!.axes!!.setTitleFont(ZAXIS, idx)
            mychart!!.setDirtyFlag(true)
        }

    /**
     * returns the underlhying Sheet Object this Chart is attached to <br></br>
     * For Internal Use
     *
     * @return Boundsheet
     */
    val sheet: Boundsheet?
        get() = mychart!!.sheet

    /**
     * return the background color of this chart's Plot Area as an int
     *
     * @return int background color constant
     * @see FormatHandle.COLOR_* constants
     */
    /**
     * sets the Plot Area background color
     *
     * @param int bg - color constant
     * @see FormatHandle.COLOR_* constants
     */
    var plotAreaBgColor: Int
        get() {
            val bg = mychart!!.plotAreaBgColor
            return FormatHandle.HexStringToColorInt(bg, 0.toShort())
        }
        set(bg) = mychart!!.setPlotAreaBgColor(bg)

    val plotAreaBgColorStr: String
        get() = mychart!!.plotAreaBgColor

    /**
     * Return an int corresponding to this ChartHandle's Chart Type for the default
     * chart <br></br>
     * To see possible Chart Types, view the public static int's in ChartHandle.
     *
     * @return int chart type
     * @see ChartHandle static Chart Type Constants
     *
     * @see ChartHandle.setChartType
     */
    /**
     * Sets the Chart type to the specified basic type (no 3d, no stacked ...) <br></br>
     * To see possible Chart Types, view the public static int's in ChartHandle.
     * <br></br>
     * Possible Chart Types: <br></br>
     * BARCHART <br></br>
     * COLCHART <br></br>
     * LINECHART <br></br>
     * PIECHART <br></br>
     * AREACHART <br></br>
     * SCATTERCHART <br></br>
     * RADARCHART <br></br>
     * SURFACECHART <br></br>
     * DOUGHNUTCHART <br></br>
     * BUBBLECHART <br></br>
     * RADARAREACHART <br></br>
     * PYRAMIDCHART <br></br>
     * CYLINDERCHART <br></br>
     * CONECHART <br></br>
     * PYRAMIDBARCHART <br></br>
     * CYLINDERBARCHART <br></br>
     * CONEBAR
     *
     * @param int chartType - representing the chart type
     */
    // no
    // specific
    // options
    var chartType: Int
        get() = mychart!!.chartType
        set(chartType) = mychart!!.setChartType(chartType, 0, EnumSet.noneOf<ChartOptions>(ChartOptions::class.java))

    /*
     * NOT IMPLEMENTED YET adjust chart cell references upon row
     * insertion or
     * deletion NOTE: Assumes we're on the correct sheet NOT
     * COMPLETELY IMPLEMENTED
     * YEt
     *
     * @param rownum
     *
     * @param shiftamt +1= insert row, -1= delete row
     *
     * public void adjustCellRefs(int rownum, int shiftamt) {
     * Vector v =
     * mychart.getAllSeries(); boolean bSeriesRows= false;
     * boolean bMod= false; for
     * (int i=0;i<v.size();i++) { Series s = (Series)v.get(i);
     * try { // SERIES Ai ai
     * = s.getSeriesValueAi(); Ptg[] p=ai.getCellRangePtgs(); //
     * should only be 1
     * ptg try { int[] loc= p[0].getIntLocation(); if
     * (loc.length==4 &&
     * loc[0]==loc[2]) bSeriesRows= true; if (shiftamt > 0) { //
     * insert row i.e.
     * shift ai location down if ((loc.length==2 &&
     * loc[0]>=rownum) ||
     * (loc[0]>=rownum || loc[2]>=rownum)) {
     * adjustAiLocation(ai,
     * p[0].getIntLocation(), shiftamt); bMod= true; } } else {
     * if ((loc.length==2
     * && loc[0]==rownum) || (loc[0]>=rownum && loc[2]<=rownum))
     * { // remove it
     * this.removeSeries(i); bMod= true; continue; } } }
     * catch(Exception e) {
     *
     * } } catch (Exception e) {
     *
     * } try { // CATEGORY Ai ai= s.getCategoryValueAi(); Ptg[]
     * p=ai.getCellRangePtgs(); try { int[] loc=
     * p[0].getIntLocation(); if (shiftamt
     * > 0) { // insert row i.e. shift ai location down if
     * ((loc.length==2 &&
     * loc[0]>=rownum) || (loc[0]>=rownum || loc[2]>=rownum)) {
     * adjustAiLocation(ai,
     * p[0].getIntLocation(), shiftamt); bMod= true; } } else {
     * if ((loc.length==2
     * && loc[0]==rownum) || (loc[0]>=rownum && loc[2]<=rownum))
     * { // remove it ????
     * this.removeSeries(i); bMod= true; } } } catch(Exception
     * e) {
     *
     * } } catch (Exception e) {
     *
     * } try { // LEGEND Ai ai= s.getLegendAi(); Ptg[]
     * p=ai.getCellRangePtgs();
     * int[] loc= p[0].getIntLocation(); if (shiftamt > 0) { //
     * insert row i.e.
     * shift ai location down if ((loc.length==2 &&
     * loc[0]>=rownum) ||
     * (loc[0]>=rownum || loc[2]>=rownum)) {
     * adjustAiLocation(ai,
     * p[0].getIntLocation(), shiftamt); bMod= true; } } else {
     * if ((loc.length==2
     * && loc[0]==rownum) || (loc[0]>=rownum && loc[2]<=rownum))
     * { // remove it
     * this.removeSeries(i); bMod= true; } } } catch(Exception
     * e) {
     *
     * } } if (bMod)// one or more series elements were modified
     * setDimensionsRecord(); // update dimensions i.e. Data
     * Range
     *
     * }
     */

    /*
     * doesn't appear to be used right now move the ai location
     * (row) according to
     * shift amount
     *
     * @param ai
     *
     * @param loc
     *
     * @param shift
     *
     * private void adjustAiLocation(Ai ai, int[] loc, int
     * shift) { String oldloc=
     * ExcelTools.formatLocation(loc); if (loc.length>2) // get
     * 2nd part of range
     * oldloc += ":" + ExcelTools.formatLocation(new
     * int[]{loc[2], loc[3]}); oldloc=
     * this.getSheet().getSheetName() + "!" + oldloc; if
     * (loc.length==2)// single
     * cell loc[0]+=shift; else { // range loc[0]+=shift;
     * loc[2]+=shift; } String
     * newloc= ExcelTools.formatLocation(loc); if (loc.length>2)
     * // get 2nd part of
     * range newloc += ":" + ExcelTools.formatLocation(new
     * int[]{loc[2], loc[3]});
     * ai.changeAiLocation(oldloc, newloc); }
     *
     */

    /**
     * Get the Chart's bytes
     *
     *
     * This is an internal method that is not useful to the end user.
     */
    val chartBytes: ByteArray?
        get() = mychart!!.chartBytes

    val serialBytes: ByteArray?
        get() = mychart!!.serialBytes

    /**
     * get the chart-type-specific options in XML form
     *
     * @return String XML
     */
    private// 0 for default chart
    val chartOptionsXML: String
        get() = mychart!!.getChartOptionsXML(0)

    /**
     * returns an XML representation of this chart
     *
     * @return String XML
     */
    // Chart Name (=Title)
    // Type
    // Plot Area Background color 20080429 KSC
    // Position
    // Chart Fonts
    // Format Chart Area
    // KSC: TODO: BORDER
    // TODO:
    // Properties
    // Source Data
    // controls shape of complex datapoints such as pyramid,
    // cylinder, cone +
    // stacked 3d bars
    // Chart Options
    // Axis Options
    // ThreeD rec opts
    val xml: String
        get() {
            val sb = StringBuffer(t(1) + "<Chart")
            sb.append(" Name=\"" + this.title + "\"")
            sb.append(" type=\"" + this.chartType + "\"")
            sb.append(" Fill=\"" + this.plotAreaBgColor + "\"")
            val coords = mychart!!.coords
            sb.append(" Left=\"" + coords[0] + "\" Top=\"" + coords[1]
                    + "\" Width=\"" + coords[2] + "\" Height=\"" + coords[3]
                    + "\"")
            sb.append(">\n")
            sb.append(t(2) + "<ChartFontRecs>" + this.chartFontRecsXML)
            sb.append("\n" + t(2) + "</ChartFontRecs>\n")
            sb.append(t(2) + "<ChartFonts" + this.chartFontsXML + "/>\n")
            sb.append(t(2) + "<FormatChartArea>\n")
            sb.append(t(3) + "<ChartBorder></ChartBorder>\n")
            sb.append(t(3) + "<ChartProperties></ChartProperties>\n")
            sb.append(t(2) + "</FormatChartArea>\n")
            sb.append(t(2) + "<SourceData>\n")
            val series = this.allChartSeriesHandles
            for (i in series.indices) {
                sb.append(t(3) + "<Series Legend=\""
                        + series[i].seriesLegendReference + "\"")
                sb.append(" Range=\"" + series[i].seriesRange + "\"")
                sb.append(" Category=\"" + series[i].categoryRange + "\"")
                if (series[i].hasBubbleSizes()) {
                    sb.append(" Bubbles=\"" + series[i].bubbleSizes + "\"")
                }
                sb.append(" TypeX=\"" + series[i].categoryDataType + "\"")
                sb.append(" TypeY=\"" + series[i].seriesDataType + "\"")
                sb.append(" Shape=\"" + series[i].shape + "\"")
                sb.append("/>\n")
            }
            sb.append(t(2) + "</SourceData>\n")
            sb.append(t(2) + "<ChartOptions")
            sb.append(this.chartOptionsXML)
            sb.append("/>\n")
            sb.append(t(2) + "<Axes>\n")
            sb.append(t(3) + "<XAxis" + this.getAxisOptionsXML(XAXIS) + "/>\n")
            sb.append(t(3) + "<YAxis" + this.getAxisOptionsXML(YAXIS) + "/>\n")
            sb.append(t(3) + "<ZAxis" + this.getAxisOptionsXML(ZAXIS) + "/>\n")
            sb.append(t(2) + "</Axes>\n")
            if (this.isThreeD) {
                sb.append(t(2) + "<ThreeD" + this.threeDXML + "/>\n")
            }

            sb.append(t(1) + "</Chart>\n")
            return sb.toString()
        }

    /**
     * return the Excel 7/OOXML-specific name for this chart
     *
     * @return String OOXML name
     */
    /**
     * set the Excel 7/OOXML-specific name for this chart
     *
     * @param String name
     */
    private var ooxmlName: String?
        get() = (mychart as OOXMLChart).ooxmlName
        set(name) {
            (mychart as OOXMLChart).ooxmlName = name
        }

    /**
     * returns the drawingml file name which defines the userShape (if any) <br></br>
     * a userShape is a drawing or shape ontop of a chart associated with this chart
     *
     * @return
     */
    val chartEmbeds: ArrayList<*>?
        get() = (mychart as OOXMLChart).chartEmbeds

    /**
     * @return true if Chart has 3D effects, false otherwise
     */
    // default chart
    val isThreeD: Boolean
        get() = mychart!!.isThreeD(0)

    /**
     * @return boolean true if Chart contains Stacked Series, false otherwise
     */
    // default chart
    val isStacked: Boolean
        get() = mychart!!.isStacked(0)

    /**
     * @return boolean true if Chart is of type 100% Stacked, false otherwise
     */
    // default chart
    val is100PercentStacked: Boolean
        get() = mychart!!.is100PercentStacked(0)

    /**
     * @return boolean true if Chart contains Clustered Bars or Columns, false
     * otherwise
     */
    // default chart
    val isClustered: Boolean
        get() = mychart!!.isClustered(0)

    /**
     * @return String ThreeD options in XML form
     */
    // 0 for default chart
    val threeDXML: String
        get() = mychart!!.getThreeDXML(0)

    /**
     * returns Chart-specific Font Records in XML form
     *
     * @return String Chart Font information in XML format
     */
    val chartFontRecsXML: String
        get() = mychart!!.chartFontRecsXML

    /**
     * Return non-axis Chart font ids in XML form
     *
     * @return String Font information in XML format
     */
    val chartFontsXML: String
        get() = mychart!!.chartFontsXML

    /**
     * @return the WorkBook Object attached to this Chart
     */
    val workBook: io.starter.formats.XLS.WorkBook?
        get() = mychart!!.workBook

    // this should be impossible
    val workSheetHandle: WorkSheetHandle
        get() {
            try {
                return this.workBookHandle.getWorkSheet(mychart!!.sheet!!.sheetNum)
            } catch (e: WorkSheetNotFoundException) {
                throw RuntimeException(e)
            }

        }

    /**
     * returns the coordinates or bounds (position, width and height) of this chart
     * in pixels
     *
     * @return short[4] bounds - left or x value, top or y value, width, height
     * @see ChartHandle.X, ChartHandle.Y, ChartHandle.WIDTH, ChartHandle.HEIGHT
     */
    // NOTE: THIS SHOULD BE RENAMED TO setCoords as Bounds and
    // Coords are very
    // distinct
    /**
     * sets the coordinates or bounds (position, width and height) of this chart in
     * pixels
     *
     * @param short[4] bounds - left or x value, top or y value, width, height
     * @see ChartHandle.X, ChartHandle.Y, ChartHandle.WIDTH, ChartHandle.HEIGHT
     */
    // NOTE: THIS SHOULD BE RENAMED TO setCoords as Bounds and
    // Coords are very
    // distinct
    var bounds: ShortArray
        get() = mychart!!.coords
        set(bounds) {
            mychart!!.coords = bounds
        }

    /**
     * returns the coordinates (position, width and height) of this chart in Excel
     * size units
     *
     * @return short[4] pixel coords - left or x value, top or y value, width,
     * height
     * @see ChartHandle.X, ChartHandle.Y, ChartHandle.WIDTH, ChartHandle.HEIGHT
     */
    /**
     * sets the coordinates (position, width and height) for this chart in Excel
     * size units
     *
     * @return short[4] pixel coords - left or x value, top or y value, width,
     * height
     * @see ChartHandle.X, ChartHandle.Y, ChartHandle.WIDTH, ChartHandle.HEIGHT
     */
    var coords: ShortArray
        get() {
            mychart!!.getMetrics(this.workBookHandle)
            return mychart!!.coords
        }
        set(coords) {
            mychart!!.coords = coords
        }

    /**
     * get the bounds of the chart using coordinates relative to row/cols and their
     * offsets
     *
     * @return short[8] bounds - COL, COLOFFSET, ROW, ROWOFFST, COL1, COLOFFSET1,
     * ROW1, ROWOFFSET1
     */
    /**
     * sets the bounds of the chart using coordinates relative to row/cols and their
     * offsets
     *
     * @param short[8] bounds - COL, COLOFFSET, ROW, ROWOFFST, COL1, COLOFFSET1, ROW1,
     * ROWOFFSET1
     */
    var relativeBounds: ShortArray?
        get() = mychart!!.bounds
        set(bounds) {
            mychart!!.bounds = bounds
        }

    /**
     * returns the offset within the column in pixels
     *
     * @return
     */
    val colOffset: Short
        get() = mychart!!.colOffset

/**
 * returs the JSON representation of this chart, based upon
 * Dojo-charting-specifics
 *
 * @return String JSON representation of the chart
 */
// necessary for parsing AXIS
// options: horizontal charts
// "switch" axes ...
// titles/labels
// bar
// axes
// are
// reversed
// ...
// Chart dimensions (width, height)
// TODO: may not be
// necessary, see usage ...
// Plot Area Background color
// it's possible to not have any series defined ...
// 20080416 KSC: Save SeriesJSON for later comparisons
// 20090729 KSC: capture
// bar
// colors or fills
// Axes + Category Labels + Grid Lines
// inputJSONObject(theChart, this.getAxis(YAXIS,
// false).getJSON(this.wbh, type,
// yMax, yMin, nSeries));
// inputJSONObject(theChart, this.getAxis(XAXIS,
// false).getJSON(this.wbh, type,
// yMax, yMin, nSeries));
// TODO: 3d Charts (z axis)
/*
             * /* Chart Fonts sb.append(t(2) + "<ChartFontRecs>" +
             * this.getChartFontRecsXML()); sb.append("\n" + t(2) +
             * "</ChartFontRecs>\n");
             * sb.append(t(2) + "<ChartFonts" + this.getChartFontsXML()
             * + "/>\n");
             *//*
             * Format Chart Area
             *//* TODO: read in legend settings */// Chart Legend
 // -1 is flag
 // for all
 // rather than
 // for a
 // specific
 // chart
 val json:String
get() {
val theChart = JSONObject()
try
{
val titles = JSONObject()
val type = this.chartType
titles.put("title", this.title)
titles.put("XAxis", if ((type != ChartConstants.BARCHART))
(this.xAxisLabel)
else
this.yAxisLabel)
titles.put("YAxis", if ((type != ChartConstants.BARCHART))
(this.yAxisLabel)
else
this.xAxisLabel)
try
{
titles.put("ZAxis", this.zAxisLabel)
}
catch (e:Exception) {
Logger.logWarn(("ChartHandle.getJSON failed getting zaxislable:" + e.toString()))
}

theChart.put("titles", titles)
val coords = mychart!!.coords
theChart.put("width", coords[ChartHandle.WIDTH].toInt())
theChart.put("height", coords[ChartHandle.HEIGHT].toInt())
theChart.put("row", mychart!!.row0)
theChart.put("col", mychart!!.col0)
var plotAreabg = this.plotAreaBgColor
if (plotAreabg == 0x4D || plotAreabg == 0x4E)
plotAreabg = FormatConstants.COLOR_WHITE
theChart.put("fill", FormatConstants.SVGCOLORSTRINGS[plotAreabg])

val jMinMax = arrayOfNulls<Double>(3)
val chartObjectJSON = this.mychart!!.chartObject
.getJSON(this.mychart!!.chartSeries, this.workBookHandle, jMinMax)
var yMax = 1.0
var yMin = 0.0
var nSeries = 0
try
{
theChart.put("Series", chartObjectJSON.getJSONArray("Series"))
mychart!!.seriesJSON = chartObjectJSON.getJSONArray("Series")
theChart.put("SeriesFills", chartObjectJSON
.getJSONArray("SeriesFills"))
}
catch (e:Exception) {}

theChart.put("type", chartObjectJSON.getJSONObject("type"))
yMin = jMinMax[0].toDouble()
yMax = jMinMax[1].toDouble()
nSeries = jMinMax[2].toInt()
try
{
theChart.put("y", mychart!!.axes!!
.getJSON(YAXIS, this.workBookHandle, type, yMax, yMin, nSeries)
.getJSONObject("y"))
theChart.put("back_grid", mychart!!.axes!!
.getJSON(YAXIS, this.workBookHandle, type, yMax, yMin, nSeries)
.getJSONObject("back_grid"))
}
catch (e:Exception) {}

try
{
theChart.put("x", mychart!!.axes!!
.getJSON(XAXIS, this.workBookHandle, type, yMax, yMin, nSeries)
.getJSONObject("x"))
theChart.put("back_grid", mychart!!.axes!!
.getJSON(ChartConstants.YAXIS, this.workBookHandle, type, yMax, yMin, nSeries)
.getJSONObject("back_grid"))
}
catch (e:Exception) {}

if (this.hasDataLegend())
{
val s = this.mychart!!.legend!!.legendPosition
val legends = this.mychart!!.getLegends(-1)
var l = ""
for (i in legends.indices)
l += legends[i] + ","
if (l.length > 0)
l = l.substring(0, l.length - 1)
theChart.put("legend", JSONObject(
"{position:" + s + ",labels:[" + l + "]}"))
}

}
catch (e:JSONException) {
Logger.logErr("Error getting Chart JSON: " + e)
}

return theChart.toString()
}

/**
 * retrieves the saved Series JSON for comparisons <br></br>
 * This is an internal method that is not useful to the end user.
 *
 * @return JSONArray
 */
    /**
 * sets the saved Series JSON <br></br>
 * This is an internal method that is not useful to the end user.
 *
 * @param JSONArray s -
 * @throws JSONException
 */
     var seriesJSON:JSONArray?
get() {
return mychart!!.seriesJSON
}
@Throws(JSONException::class)
set(s) {
mychart!!.seriesJSON = s
}

/**
 * retrieves current series and axis scale info in JSONObject form used upon
 * chart updating <br></br>
 * This is an internal method that is not useful to the end user.
 *
 * @return JSONObject series and axis info
 */
    // Retrieve series data + yMin yMax, nSeries
 // Series Data
 // 20080516 KSC: See above JSONObject chartObjectJSON=
 // ((GenericChartObject)this.mychart.getChartObject()).getJSON(
 // this.getAllChartSeriesHandles(), this.getCategories()[0],
 // this.wbh, minMax);
 // Retrieve Axis Scale info
 // necessary for parsing AXIS
 // options: horizontal charts
 // "switch" axes ...
 // Axes + Category Labels + Grid Lines
 /*
             * KSC: TAKE OUT JSON STUFF FOR NOW; WILL REFACTOR LATER try
             * {
             * inputJSONObject(retJSON,
             * mychart.getAxes().getMinMaxJSON(YAXIS, this.wbh,
             * type, yMax, yMin, nSeries)); } catch (Exception e) { }
             * try {
             * inputJSONObject(retJSON,
             * mychart.getAxes().getMinMaxJSON(XAXIS, this.wbh,
             * type, yMax, yMin, nSeries)); } catch (Exception e) { }
             */// TODO: 3d Charts (z axis)
 val currentSeries:JSONObject
get() {
val retJSON = JSONObject()
val jMinMax = arrayOfNulls<Double>(3)
try
{
val chartObjectJSON = this.mychart!!.chartObject
.getJSON(this.mychart!!.chartSeries, this.workBookHandle, jMinMax)

try
{
retJSON.put("Series", chartObjectJSON.getJSONArray("Series"))
}
catch (e:Exception) {
Logger.logWarn(("ChartHandle.getCurrentSeries problem:" + e.toString()))
}

var yMax = 0.0
var yMin = 0.0
var nSeries = 0
yMin = jMinMax[0].toDouble()
yMax = jMinMax[1].toDouble()
nSeries = jMinMax[2].toInt()

val type = this.chartType
}
catch (e:JSONException) {
Logger.logErr(("ChartHandle.getCurrentSeries: Error getting Series JSON: " + e))
}

return retJSON
}

/**
 * returns a JSON representation of all Series Data (Legend, Categogies, Series
 * Values) for the chart <br></br>
 * This is an internal method that is not useful to the end user.
 *
 * @return String JSON representation
 */
    // 1 per chart
 val allSeriesDataJSON:String
get() {
val s = JSONArray()
val series = this.allChartSeriesHandles
try
{
for (i in series.indices)
{
val ser = JSONObject()
ser.put("l", series[i].seriesLegendReference)
ser.put("v", series[i].seriesRange)
ser.put("b", series[i].bubbleSizes)
if (i == 0)
ser.put("c", series[i].categoryRange)
s.put(ser)
}
}
catch (e:JSONException) {
Logger.logErr("ChartHandle.getAllSeriesDataJSON: " + e)
}

return s.toString()
}

/**
 * Take current Chart object and return the SVG code necessary to define it.
 */
    /**
 * TODO: Less Common Charts: STOCK RADAR SURFACE COLUMN- 3D, CONE, CYLINDER,
 * PYRAMID BAR- 3D, CONE, CYLINDER, PYRAMID 3D PIE 3D LINE 3D AREA
 *
 *
 * LINE CHART APPEARS THAT STARTS AND ENDS A BIT TOO EARLY ***************** Z
 * Axis
 *
 *
 * CHART OPTIONS: STACKED CLUSTERED
 */
     val svg:String
get() {
return getSVG(1.0)
}

/**
 * returns the svg for javascript for highlight and restore
 *
 * @return
 */
    protected// rgb('+
 // red
 // +','+
 // green+','+blue+')');");
 // svg.append("evt.target.setAttributeNS(null,'stroke-color','white');");
 // rgb('+
 // red
 // +','+
 // green+','+blue+')');");
 // svg.append("try{parent.parent.uiWindowing.getActiveSheet().book.handleMouseClick(evt);}catch(x){;}");
 val javaScript:String
get() {
val svg = StringBuffer()
svg.append("<script type='text/ecmascript'>")
svg.append("  <![CDATA[")

svg.append("try{var grid = parent.parent.uiWindowing.getActiveDoc().getContent().contentWindow;")
svg.append("var selection = new grid.cellSelectionResizable();}catch(x){;}")

svg.append("function highLight(evt) {")
svg.append("this.bgc = evt.target.getAttributeNS(null, 'fill');")
svg.append("evt.target.setAttributeNS(null,'fill','gold');")
svg.append("evt.target.setAttributeNS(null,'stroke-width','2');")
svg.append("}")

svg.append("function restore(evt) {")
svg.append("evt.target.setAttributeNS(null,'fill',this.bgc);")
svg.append("evt.target.setAttributeNS(null,'stroke-width','1');")
svg.append("}")

svg.append("function handleClick(evt) {")
svg.append("try{parent.chart.handleClick(evt);}catch(e){;}")
svg.append("}")

svg.append("\r\n")
svg.append("function showRange(range) {")
svg.append("try{selection.select(new grid.cellRange(range.toString())).show();}catch(x){;}")
svg.append("}")

svg.append("\r\n")
svg.append("function hideRange() {")
svg.append("try{selection.clear();}catch(x){;}")
svg.append("}")

svg.append("]]>")
svg.append("</script>")
return svg.toString()
}

init{
workBookHandle = wb
if (mychart!!.workBook == null)
 // TODO: WHY IS THIS NULL????
            mychart!!.workBook = wb.workBook
}// super();

/**
 * returns the string representation of this ChartHandle
 */
    public override fun toString():String {
return mychart!!.title
}

/**
 * Returns an ordered array of strings representing all the category ranges in
 * the chart. <br></br>
 * This vector corresponds to the getSeries() method so will often contain
 * duplicates, as while the series data changes frequently, category data is the
 * same throughout the chart.
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @return String[] each item being a Cell Range representing the Category Data
 * @see ChartHandle.getSeries
 */
    private fun getCategories(nChart:Int):Array<String> {
return mychart!!.getCategories(nChart)
}

/**
 * Returns an array of ChartSeriesHandle Objects for the desired chart, one for
 * each bar, line or wedge of data. <br></br>
 * A chart number of 0 means the default chart, 1-9 indicate series for overlay
 * charts <br></br>
 * NOTE: using this method returns the series for the desired chart ONLY <br></br>
 * You MUST use the corresponding removeSeries(index, nChart) when removing
 * series to properly match the series index. Otherwise a mismatch will occur.
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @return ChartSeriesHandle[] Array of ChartSeriesHandle Objects representing
 * Chart Series Data (Series and Categories)
 * @see ChartSeriesHandle
 */
     fun getAllChartSeriesHandles(nChart:Int):Array<ChartSeriesHandle> {
val v = mychart!!.getAllSeries(nChart)
val csh = arrayOfNulls<ChartSeriesHandle>(v.size)
for (i in v.indices)
{
val s = v.get(i) as Series
csh[i] = ChartSeriesHandle(s, this.workBookHandle)
}
return csh
}

/**
 * Get the ChartSeriesHandle representing Chart Series Data (Series and
 * Categories) for the specified Series range
 *
 * @param String seriesRange - For example, "Sheet1!A12:A21"
 * @return ChartSeriesHandle
 * @see ChartSeriesHandle
 */
     fun getChartSeriesHandle(seriesRange:String):ChartSeriesHandle? {
val series = this.allChartSeriesHandles
for (i in series.indices)
{
val sr = series[i].seriesRange
if (seriesRange.equals(sr, ignoreCase = true))
{
return series[i]
}
}
return null
}

/**
 * Get the ChartSeriesHandle representing Chart Series Data (Series and
 * Categories) for the specified Series index
 *
 * @param int idx - the index (0 based) of the series
 * @return ChartSeriesHandle
 * @see ChartSeriesHandle
 */
     fun getChartSeriesHandle(idx:Int):ChartSeriesHandle? {
val series = this.allChartSeriesHandles
if (series.size >= idx)
return series[idx]
return null
}

/**
 * Get the ChartSeriesHandle representing Chart Series Data (Series and
 * Categories) for the Series specified by label (legend)
 *
 * @param String legend - label for the desired series
 * @return ChartSeriesHandle
 * @see ChartSeriesHandle
 */
     fun getChartSeriesHandleByName(legend:String):ChartSeriesHandle {
val s = mychart!!.getSeries(legend, -1) // -1 is flag for all rather
 // than for a specific chart
        return ChartSeriesHandle(s, this.workBookHandle)
}

/**
 * sets or removes the axis title
 *
 * @param axisType one of: XAXIS, YAXIS, ZAXIS
 * @param ttl      String new title or null to remove
 */
     fun setAxisTitle(axisType:Int, ttl:String) {
mychart!!.axes!!.setTitle(axisType, ttl)
}

/**
 * Sets the automatic scale option on or off for the desired Value axis
 *
 *
 * The Value axis contains numbers rather than labels, and is normally the Y
 * axis, but Scatter and Bubble charts may have a value axis on the X Axis as
 * well
 *
 *
 * Automatic Scaling automatically sets the scale maximum, minimum and tick
 * units upon data changes, and is the default setting for charts
 *
 * @param int     axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS,
 * ChartHandle.ZAXIS
 * @param boolean b - true if set Automatic scaling on, false otherwise
 * @see setAxisAutomaticScale
 */
     fun setAxisAutomaticScale(axisType:Int, b:Boolean) {
mychart!!.axes!!.setAxisAutomaticScale(axisType, b)
mychart!!.setDirtyFlag(true)
}

/**
 * Returns the minimum scale value of the the desired Value axis
 *
 *
 * The Value axis contains numbers rather than labels, and is normally the Y
 * axis, but Scatter and Bubble charts may have a value axis on the X Axis as
 * well
 *
 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS,
 * ChartHandle.ZAXIS
 * @return int Miminum Scale value for the desired axis
 * @see getAxisMinScale
 */
     fun getAxisMinScale(axisType:Int):Double {
val minmax = mychart!!.getMinMax(this.workBookHandle)
return mychart!!.axes!!.getMinMax(minmax[0], minmax[1], axisType)[0]
}

/**
 * Returns the maximum scale value of the desired Value axis
 *
 *
 * The Value axis contains numbers rather than labels, and is normally the Y
 * axis, but Scatter and Bubble charts may have a value axis on the X Axis as
 * well
 *
 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS,
 * ChartHandle.ZAXIS
 * @return int - Maximum Scale value for the desired axis
 * @see getAxisMaxScale
 */
     fun getAxisMaxScale(axisType:Int):Double {
val minmax = mychart!!.getMinMax(this.workBookHandle) // -1 is flag for all
 // rather than for a
        // specific chart
        return mychart!!.axes!!.getMinMax(minmax[0], minmax[1], axisType)[1]
}

/**
 * Returns the major tick unit of the desired Value axis
 *
 *
 * The Value axis contains numbers rather than labels, and is normally the Y
 * axis, but Scatter and Bubble charts may have a value axis on the X Axis as
 * well
 *
 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS,
 * ChartHandle.ZAXIS
 * @return int major tick unit
 * @see getAxisMajorUnit
 */
     fun getAxisMajorUnit(axisType:Int):Int {
return mychart!!.axes!!.getAxisMajorUnit(axisType)
}

/**
 * Returns the minor tick unit of the desired Value axis
 *
 *
 * The Value axis contains numbers rather than labels, and is normally the Y
 * axis, but Scatter and Bubble charts may have a value axis on the X Axis as
 * well
 *
 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS,
 * ChartHandle.ZAXIS
 * @return int - minor tick unit of the desired axis
 * @see getAxisMinorUnit
 */
     fun getAxisMinorUnit(axisType:Int):Int {
return mychart!!.axes!!.getAxisMinorUnit(axisType)
}

/**
 * Sets the maximum scale value of the desired Value axis
 *
 *
 * The Value axis contains numbers rather than labels, and is normally the Y
 * axis, but Scatter and Bubble charts may have a value axis on the X Axis as
 * well
 *
 *
 * Note: The default scale setting for charts is known as Automatic Scaling <br></br>
 * When data changes, the chart automatically adjusts the scale as necessary
 * <br></br>
 * Setting the scale manually (either Minimum or Maximum Value) removes
 * Automatic Scaling
 *
 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS,
 * ChartHandle.ZAXIS
 * @param int MaxValue - desired maximum value of the desired axis
 * @see setAxisMax
 */
     fun setAxisMax(axisType:Int, MaxValue:Int) {
mychart!!.axes!!.setAxisMax(axisType, MaxValue)
mychart!!.setDirtyFlag(true)
}

/**
 * Sets the minimum scale value of the desired Value axis
 *
 *
 * The Value axis contains numbers rather than labels, and is normally the Y
 * axis, but Scatter and Bubble charts may have a value axis on the X Axis as
 * well
 *
 *
 * Note: The default setting for charts is known as Automatic Scaling <br></br>
 * When data values change, the chart automatically adjusts the scale as
 * necessary <br></br>
 * Setting the scale manually (either Minimum or Maximum Value) removes
 * Automatic Scaling
 *
 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS,
 * ChartHandle.ZAXIS
 * @param int MinValue - the desired Minimum scale value
 * @see setAxisMin
 */
     fun setAxisMin(axisType:Int, MinValue:Int) {
mychart!!.axes!!.setAxisMin(axisType, MinValue)
mychart!!.setDirtyFlag(true)
}

/**
 * Returns true if the desired Value axis is set to automatic scale
 *
 *
 * The Value axis contains numbers rather than labels, and is normally the Y
 * axis, but Scatter and Bubble charts may have a value axis on the X Axis as
 * well
 *
 *
 * Note: The default setting for charts is known as Automatic Scaling <br></br>
 * When data changes, the chart automatically adjusts the scale as necessary
 * <br></br>
 * Setting the scale manually (either Minimum or Maximum Value) removes
 * Automatic Scaling
 *
 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS,
 * ChartHandle.ZAXIS
 * @return boolean true if Automatic Scaling is turned on
 * @see getAxisAutomaticScale
 */
     fun getAxisAutomaticScale(axisType:Int):Boolean {
return mychart!!.axes!!.getAxisAutomaticScale(axisType)
}

/**
 * Sets the maximum value of the Y Axis (Value Axis) Scale
 *
 *
 * Note: The default scale setting for charts is known as Automatic Scaling <br></br>
 * When data changes, the chart automatically adjusts the scale as necessary
 * <br></br>
 * Setting the scale manually (either Minimum or Maximum Value) removes
 * Automatic Scaling
 *
 * @param int MaxValue - the desired maximum scale value
 * @see ChartHandle.setAxisMax
 */
     fun setAxisMax(MaxValue:Int) {
mychart!!.axes!!.setAxisMax(MaxValue)
mychart!!.setDirtyFlag(true)
}

/**
 * Sets the minimum value of the Y Axis (Value Axis) Scale
 *
 *
 * Note: The default setting for charts is known as Automatic Scaling <br></br>
 * When data changes, the chart automatically adjusts the scale as necessary
 * <br></br>
 * Setting the scale manually (either Minimum or Maximum Value) removes
 * Automatic Scaling
 *
 * @param int MinValue - the desired minimum scale value
 * @see ChartHandle.setAxisMin
 */
     fun setAxisMin(MinValue:Int) {
mychart!!.axes!!.setAxisMin(MinValue)
mychart!!.setDirtyFlag(true)
}

/**
 * sets an option for this axis
 *
 * @param axisType one of: XAXIS, YAXIS, ZAXIS
 * @param op       option name; one of: CatCross, LabelCross, Marks, CrossBetween,
 * CrossMax, MajorGridLines, AddArea, AreaFg, AreaBg or Linked Text
 * Display options: Label, ShowKey, ShowValue, ShowLabelPct, ShowPct,
 * ShowCatLabel, ShowBubbleSizes, TextRotation, Font
 * @param val      option value
 */
     fun setAxisOption(axisType:Int, op:String, `val`:String) {
mychart!!.axes!!.setChartOption(axisType, op, `val`)

}

/**
 * set the font for the chart title
 *
 * @param String  name - font name
 * @param int     height - font height in 1/20 point units
 * @param boolean bold - true if bold
 * @param boolean italic - true if italic
 * @param boolean underline - true if underlined
 */
     fun setTitleFont(name:String, height:Int, bold:Boolean, italic:Boolean, underline:Boolean) {
val f = Font(name, 200, height)
f.bold = bold
f.italic = italic
f.underlined = underline
var idx = workBookHandle.workBook!!.getFontIdx(f)
if (idx == -1)
{ // font doesn't exist yet, add to streamer
f.idx = -1 // flag to insert
idx = workBookHandle.workBook!!.insertFont(f) + 1
}
mychart!!.setTitleFont(idx)
}

/**
 * set the font for all axes on the chart
 *
 * @param String  name - font name
 * @param int     height - font height in 1/20 point units
 * @param boolean bold - true if bold
 * @param boolean italic - true if italic
 * @param boolean underline - true if underlined
 */
     fun setAxisFont(name:String, height:Int, bold:Boolean, italic:Boolean, underline:Boolean) {
val f = Font(name, 200, height)
f.bold = bold
f.italic = italic
f.underlined = underline
var idx = workBookHandle.workBook!!.getFontIdx(f)
if (idx == -1)
{ // font doesn't exist yet, add to streamer
f.idx = -1 // flag to insert
idx = workBookHandle.workBook!!.insertFont(f) + 1
}
mychart!!.axes!!.setTitleFont(XAXIS, idx)
mychart!!.axes!!.setTitleFont(YAXIS, idx)
mychart!!.axes!!.setTitleFont(ZAXIS, idx)
mychart!!.setDirtyFlag(true)
}

/**
 * resets all fonts in the chart to the default font of the workbook
 */
     fun resetFonts() {
mychart!!.resetFonts()
}

/**
 * Change the value of a Chart object. <br></br>
 * **NOTE: THIS HAS NOT BEEN 100% IMPLEMENTED YET** <br></br>
 * You can use this method to change:
 *
 * <br></br>
 * - the Title of the Chart <br></br>
 * - the Text Labels of Categories and Values (X and Y)
 *
 * <br></br>
 * eg:
 *
 * <br></br>
 * To change the value of the Chart title <br></br>
 * chart.changeObjectValue("Template Chart Title", "Widget Sales By Quarter");
 *
 * <br></br>
 * To change the text label of the categories <br></br>
 * chart.changeObjectValue("Category X", "Fiscal Year");
 *
 * <br></br>
 * To change the text label of the values <br></br>
 * chart.changeObjectValue("Value Y", "Sales in US$");
 *
 * @param String originalval - One of: "Template Chart Title", "Category X" or
 * "Value Y"
 * @param Sring  newval - the new setting
 * @return whether the change was successful
 */
     fun changeTextValue(originalval:String, newval:String):Boolean {
 /*
         * KSC: TODO: Refactor *** for(int
         * x=0;x<mychart.aivals.size();x++){ Ai ser =
         * (Ai)mychart.aivals.get(x);
         * if(ser.toString().equalsIgnoreCase(originalval)){
         * return ser.setText(newval); } }
         */
        return mychart!!.changeTextValue(originalval, newval)
 // return false;
    }

/**
 * Sets the location lock on the Cell Reference at the specified location
 *
 *
 * Used to prevent updating of the Cell Reference when Cells are moved.
 *
 * @param location of the Cell Reference to be locked/unlocked
 * @param lock     status setting
 * @return boolean whether the Cell Reference was found and modified
 */
    private fun setLocationPolicy(loc:String, l:Int):Boolean {
Logger.logErr("ChartHandle.setLocationPolicy is broken")

 /*
         * TODO: Refactor for(int x=0;x<mychart.aivals.size();x++){
         * Ai ser =
         * (Ai)mychart.aivals.get(x);
         * if(ser.setLocationPolicy(loc,l)){ return true; } }
         */
        return false
}

/**
 * Chart Options. CLUSTERED, bar, col charts only STACKED, PERCENTSTACKED, 100%
 * stacked THREED, 3d Effect EXPLODED, Pie, Donut HASLINES, Scatter, Line charts
 * ... WIREFRAME, Surface DROPLINES, DOWNBARS, line, stock UPDOWNBARS, line,
 * stock SERLINES, bar, ofpie <br></br>
 * Use these chart options when creating new charts <br></br>
 * A chart may have multiple chart options e.g. 3D Exploded pie chart
 *
 * @see ChartHandle.createNewChart
 */
    // need: hasMarkers ****
     enum class ChartOptions {
CLUSTERED,
/**
 * bar, col charts only
 */
        STACKED,
PERCENTSTACKED,
/**
 * 100% stacked
 */
        THREED,
/**
 * 3d Effect
 */
        EXPLODED,
/**
 * Pie, Donut
 */
        HASLINES,
/**
 * Scatter, Line
 */
        SMOOTHLINES,
/**
 * Scatter, Line, Radar
 */
        WIREFRAME,
/**
 * Surface
 */
        DROPLINES,
/**
 * line, area, stock charts
 */
        UPDOWNBARS,
/**
 * line, stock
 */
        SERLINES,
/**
 * bar, ofpie
 */
        HILOWLINES,
/**
 * line, stock charts
 */
        FILLED /** radar  */
}

/**
 * Sets the Chart type to the specified type (no 3d, no stacked ...) <br></br>
 * To see possible Chart Types, view the public static int's in ChartHandle.
 * <br></br>
 * Possible Chart Types: <br></br>
 * BARCHART <br></br>
 * COLCHART <br></br>
 * LINECHART <br></br>
 * PIECHART <br></br>
 * AREACHART <br></br>
 * SCATTERCHART <br></br>
 * RADARCHART <br></br>
 * SURFACECHART <br></br>
 * DOUGHNUTCHART <br></br>
 * BUBBLECHART <br></br>
 * RADARAREACHART <br></br>
 * PYRAMIDCHART <br></br>
 * CYLINDERCHART <br></br>
 * CONECHART <br></br>
 * PYRAMIDBARCHART <br></br>
 * CYLINDERBARCHART <br></br>
 * CONEBARCHART
 *
 * @param int                   chartType - representing the chart type
 * @param nChart                - 0 (default) or 1-9 for complex overlay charts
 * @param EnumSet<ChartOptions> 0 or more chart options (Such as Stacked, Exploded ...)
 * @see ChartHandle.ChartOptions
</ChartOptions> */
     fun setChartType(chartType:Int, nChart:Int, options:EnumSet<ChartOptions>) {
mychart!!.setChartType(chartType, nChart, options)
}

/**
 * Sets the basic chart type (no 3d, stacked...) for multiple or overlay Charts.
 * <br></br>
 * You can specify the drawing order of the Chart, where 0 is the default chart,
 * and 1-9 are overlay charts. <br></br>
 * The default chart (chart 0) is always present; however, using this method,
 * you can create a new overlay chart (up to 9 maximum). <br></br>
 * NOTE: The chart number must be **unique** and **in order** <br></br>
 * If the desired chart number is not present in the chart, a new overlay chart
 * will be created. <br></br>
 * **To set explicit chart options, @see setChartType(chartType, nChart, is3d,
 * isStacked, is100PercentStacked)** <br></br>
 * <br></br>
 * To see possible Chart Types, view the public static int's in ChartHandle.
 * <br></br>
 * Possible Chart Types: <br></br>
 * BARCHART <br></br>
 * COLCHART <br></br>
 * LINECHART <br></br>
 * PIECHART <br></br>
 * AREACHART <br></br>
 * SCATTERCHART <br></br>
 * RADARCHART <br></br>
 * SURFACECHART <br></br>
 * DOUGHNUTCHART <br></br>
 * BUBBLECHART <br></br>
 * RADARAREACHART <br></br>
 * PYRAMIDCHART <br></br>
 * CYLINDERCHART <br></br>
 * CONECHART <br></br>
 * PYRAMIDBARCHART <br></br>
 * CYLINDERBARCHART <br></br>
 * CONEBARCHART
 *
 * @param int       chartType - representing the chart type
 * @param chartType
 * @param nChart    number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 */
     fun setChartType(chartType:Int, nChart:Int) {
mychart!!.setChartType(chartType, nChart, EnumSet
.noneOf<ChartOptions>(ChartOptions::class.java)) // no specific options
}

/**
 * Return an int corresponding to this ChartHandle's Chart Type for the
 * specified chart <br></br>
 * To see possible Chart Types, view the public static int's in ChartHandle.
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @return int chart type
 * @see ChartHandle static Chart Type Constants
 *
 * @see ChartHandle.setChartType
 */
     fun getChartType(nChart:Int):Int {
return mychart!!.getChartType(nChart)
}

/**
 * Sets the location lock on the Cell Reference at the specified location
 *
 * Used to prevent updating of the Cell Reference when Cells are moved.
 *
 * @param location
 * of the Cell Reference to be locked/unlocked
 * @param lock
 * status setting
 * @return boolean whether the Cell Reference was found and modified
 */
    /*
     * TODO: NEEDED?? public boolean setLocationLocked(String
     * loc, boolean l){ int x
     * = Ptg.PTG_LOCATION_POLICY_UNLOCKED; if(l)x =
     * Ptg.PTG_LOCATION_POLICY_LOCKED;
     * return setLocationPolicy(loc, x); }
     */

    /**
 * Change the Cell Range referenced by one of the Series (a bar, line or wedge
 * of data) in the Chart. <br></br>
 * <br></br>
 * For Example, if the data values for one of the Series in the Chart are
 * obtained from the range Sheet1!A1:A10 and we want to add 5 more values to
 * that Series, use:
 *
 * <br></br>
 * boolean changedOK =
 * charthandle.changeSeriesRange("Sheet1!A1:A10","Sheet1!A1:A15"); <br></br>
 * <br></br>
 * Please keep in mind this is only the data range; it does not include labels
 * that may have been automatically created when you chose the chart range. <br></br>
 * <br></br>
 * To illustrate this, if A1 = "label" A2 = "data" A3 = "data", and we want to
 * add 2 more data points, we would use: changeSeriesRange("Sheet1!A2:A3",
 * "Sheet1!A2:A5"); <br></br>
 * <br></br>
 * Series are always expressed as one single line of data. If your chart
 * encompasses a range of rows and columns you will need to modify each of the
 * series in the chart handle. To determine the series that already exist in
 * your chart, utilize the String[] getSeries() method.
 *
 * @param String originalrange - the original Series (bar, line or wedge of data)
 * to alter
 * @param String newrange -the new data range
 * @return whether the change was successful
 */
     fun changeSeriesRange(originalrange:String, newrange:String):Boolean {
return mychart!!.changeSeriesRange(originalrange, newrange)
}

/**
 * Change the Cell Range representing the Categories in the Chart. <br></br>
 * Categories usually appear on the X Axis and are textual, not numeric <br></br>
 * For example: the Category values in the Chart are obtained from the range
 * Sheet1!A1:A10 and we want to add 5 more categories to the chart: <br></br>
 * boolean changedOK =
 * chart.changeCategoryRange("Sheet1!A1:A10","Sheet1!A1:A15"); <br></br>
 * Note that Category Range is the same for each Series (bar, line or wedge of
 * data) <br></br>
 * i.e. there is only one Category Range for the Chart, but there may be many
 * Series Ranges
 *
 * @param String originalrange - Original Category Range
 * @param String newrange - New Category Range
 * @return true if the change was successful
 */
     fun changeCategoryRange(originalrange:String, newrange:String):Boolean {
return mychart!!.changeCategoryRange(originalrange, newrange)
}

/**
 * Changes or adds a Series to the chart via Series Index. Each bar, line or
 * wedge in a chart represents a Series. <br></br>
 * If the Series index is greater than the number of series already present in
 * the chart, the series will be added to the end. <br></br>
 * Otherwise the Series at the index position will be altered. <br></br>
 * This method allows altering of every aspect of the Series: Data (Series)
 * Range, Legend Cell Address, Category Range and/or Bubble Range.
 *
 * @param int    index - the series index. If greater than the number of series
 * already present in the chart, the series will be added to the end
 * @param String legendCell - String representation of Legend Cell Address
 * @param String categoryRange - String representation of Category Range (should be
 * same for all series)
 * @param String seriesRange - String representation of the Series Data Range for
 * this series
 * @param String bubbleRange - String representation of Bubble Range (representing
 * bubble sizes), if bubble chart. null if not
 * @return a ChartSeriesHandle representing the new or altered Series
 * @throws CellNotFoundException
 */
    @Throws(CellNotFoundException::class)
 fun setSeries(index:Int, legendCell:String?, categoryRange:String, seriesRange:String, bubbleRange:String):ChartSeriesHandle {
var legendText:String? = ""
try
{
var ICell:CellHandle? = null
if (legendCell != null && legendCell != "")
{
 // 20070707 KSC: allow addition of new cell ranges for
                // legendCell (see
                // OpenXLS.handleChartElement)
                try
{
ICell = workBookHandle.getCell(legendCell!!)
}
catch (c:CellNotFoundException) {
val shtpos = legendCell!!.indexOf("!")
if (shtpos > 0)
{
val sheetstr = legendCell!!.substring(0, shtpos)
val sht = workBookHandle.getWorkSheet(sheetstr)
val celstr = legendCell!!.substring(shtpos + 1)
ICell = sht!!.add("", celstr)
}
}

legendText = ICell!!.stringVal
}
return setSeries(index, legendCell, legendText, categoryRange, seriesRange, bubbleRange)
}
catch (e:WorkSheetNotFoundException) {
throw CellNotFoundException(
("Error locating cell for adding series range: " + legendCell!!))
}

}

/**
 * Changes or adds a Series to the desired Chart (either default or overlay) via
 * Series Index. Each bar, line or wedge in a chart represents a Series. <br></br>
 * If the Series index is greater than the number of series already present in
 * the chart, the series will be added to the end. <br></br>
 * Otherwise the Series at the index position will be altered. <br></br>
 * This method allows altering of every aspect of the Series: Data (Series)
 * Range, Legend Text, Legend Cell Address, Category Range and/or Bubble Range.
 *
 * @param int    index - the series index. If greater than the number of series
 * already present in the chart, the series will be added to the end
 * @param String legendCell - String representation of Legend Cell Address
 * @param String legendText - String Legend text
 * @param String categoryRange - String representation of Category Range (should be
 * same for all series)
 * @param String seriesRange - String representation of the Series Data Range for
 * this series
 * @param String bubbleRange - String representation of Bubble Range (representing
 * bubble sizes), if bubble chart. null if not
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @return a ChartSeriesHandle representing the new or altered Series
 * @throws CellNotFoundException
 */
    @Throws(CellNotFoundException::class)
@JvmOverloads  fun setSeries(index:Int, legendCell:String?, legendText:String, categoryRange:String, seriesRange:String, bubbleRange:String, nChart:Int = 0):ChartSeriesHandle {
 // if (index < mychart.getAllSeries(nChart).size() && index
        // >= 0) {
        try
{
val s = mychart!!.getAllSeries(nChart).get(index) as Series
val csh = ChartSeriesHandle(s, this.workBookHandle)
csh.setSeries(legendCell, categoryRange, seriesRange, bubbleRange)
setDimensionsRecord()
return csh
}
catch (ae:ArrayIndexOutOfBoundsException) { // not found - add
return addSeriesRange(legendCell, legendText, categoryRange, seriesRange, bubbleRange, nChart)
}

}

/**
 * Adds a new Series to the chart. Each bar, line or wedge in a chart represents
 * a Series.
 *
 * @param String legendAddress - The cell address defining the legend for the
 * series
 * @param Srring legendText - Text of the legend
 * @param String categoryRange - Cell Range defining the category (normally will be
 * the same range for all series)
 * @param String seriesRange - Cell range defining the data points of the series
 * @param String bubbleRange - Cell range defining the bubble sizes for this series
 * (bubble charts only)
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @return ChartSeriesHandle representing the new series
 * @throws CellNotFoundException
 */
    @Throws(CellNotFoundException::class)
private fun addSeriesRange(legendAddress:String?, legendText:String, categoryRange:String, seriesRange:String, bubbleRange:String?, nChart:Int = 0):ChartSeriesHandle {
var s:Series? = null
if (bubbleRange == null || bubbleRange == "")
s = mychart!!
.addSeries(seriesRange, categoryRange, "", legendAddress, legendText, nChart)
else
s = mychart!!
.addSeries(seriesRange, categoryRange, bubbleRange, legendAddress, legendText, nChart)
if (nChart > 0)
{ // must update SeriesList record for overlay charts
 // TODO: FINISH
        }
setDimensionsRecord()
return ChartSeriesHandle(s, workBookHandle)
}

/**
 * Adds a new Series to the chart. Each bar, line or wedge in a chart represents
 * a Series. <br></br>
 * An Example of adding multiple series to a chart:
 *
 *
 * ChartHandle.addSeriesRange("Sheet1!A3", "Sheet1!B1:E1", "Sheet1:B3:E3",
 * null); <br></br>
 * ChartHandle.addSeriesRange("Sheet1!A4", "Sheet1!B1:E1", "Sheet1:B4:E4",
 * null); <br></br>
 * ChartHandle.addSeriesRange("Sheet1!A5", "Sheet1!B1:E1", "Sheet1:B5:E5",
 * null); <br></br>
 * etc...
 *
 *
 * Note that the category does not change, it is usually constant through
 * series. <br></br>
 * Also note that the example above is for a non-bubble-type chart.
 *
 * @param String legendCell - Cell reference for the legend cell (e.g. Sheet1!A1)
 * @param String categoryRange - Category Cell range (e.g. Sheet1!B1:B1);
 * @param String seriesRange - Series Data range (e.g. Sheet1!B3:E3);
 * @param String bubbleRange - Cell Range representing Bubble sizes (e.g.
 * Sheet1!A2:A5); or null if chart is not of type Bubble.
 * @return ChartSeriesHandle referencing the newly added series
 */
    @Throws(CellNotFoundException::class)
 fun addSeriesRange(legendCell:String, categoryRange:String, seriesRange:String, bubbleRange:String):ChartSeriesHandle {
return this
.addSeriesRange(legendCell, categoryRange, seriesRange, bubbleRange, 0) // target
 // default
        // chart
    }

/**
 * Adds a new Series to the chart. Each bar, line or wedge in a chart represents
 * a Series. <br></br>
 * An Example of adding multiple series to a chart:
 *
 *
 * ChartHandle.addSeriesRange("Sheet1!A3", "Sheet1!B1:E1", "Sheet1:B3:E3",
 * null); <br></br>
 * ChartHandle.addSeriesRange("Sheet1!A4", "Sheet1!B1:E1", "Sheet1:B4:E4",
 * null); <br></br>
 * ChartHandle.addSeriesRange("Sheet1!A5", "Sheet1!B1:E1", "Sheet1:B5:E5",
 * null); <br></br>
 * etc...
 *
 *
 * Note that the category does not change, it is usually constant through
 * series. <br></br>
 * Also note that the example above is for a non-bubble-type chart.
 *
 * @param String legendCell - Cell reference for the legend cell (e.g. Sheet1!A1)
 * @param String categoryRange - Category Cell range (e.g. Sheet1!B1:B1);
 * @param String seriesRange - Series Data range (e.g. Sheet1!B3:E3);
 * @param String bubbleRange - Cell Range representing Bubble sizes (e.g.
 * Sheet1!A2:A5); or null if chart is not of type Bubble.
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @return ChartSeriesHandle referencing the newly added series
 */
    @Throws(CellNotFoundException::class)
 fun addSeriesRange(legendCell:String?, categoryRange:String, seriesRange:String, bubbleRange:String?, nChart:Int):ChartSeriesHandle {
var legendText:String? = ""
var legendAddr = ""
try
{
var ICell:CellHandle? = null
if (legendCell != null && legendCell != "")
{
try
{
ICell = workBookHandle.getCell(legendCell!!)
if (legendCell!!.indexOf("!") == -1)
legendAddr = (ICell!!.workSheetName + "!"
+ ICell!!.cellAddress)
else
legendAddr = legendCell
}
catch (c:CellNotFoundException) {
val shtpos = legendCell!!.indexOf("!")
if (shtpos > 0)
{
val sheetstr = legendCell!!.substring(0, shtpos)
val sht = workBookHandle.getWorkSheet(sheetstr)
val celstr = legendCell!!.substring(shtpos + 1)
ICell = sht!!.add("", celstr) // TODO: Why is this being
 // added?
                        legendAddr = celstr
}
}

if (ICell != null)
legendText = ICell!!.stringVal
else
legendText = legendCell
}
var s:Series? = null
if (bubbleRange == null)
s = mychart!!
.addSeries(seriesRange, categoryRange, "", legendAddr, legendText, nChart)
else
s = mychart!!
.addSeries(seriesRange, categoryRange, bubbleRange, legendAddr, legendText, nChart)
setDimensionsRecord() // update chart DIMENSIONS record upon update
 // of series
            return ChartSeriesHandle(s, workBookHandle)
}
catch (e:WorkSheetNotFoundException) {
throw CellNotFoundException(
("Error locating cell for adding series range: " + legendCell!!))
}

}

/**
 * Adds a new Series to the chart via CellHandles and CellRange Objects. Each
 * bar, line or wedge in a chart represents a Series.
 *
 * @param CellHandle legendCell - references the legend cell for this series
 * @param CellRange  categoryRange - The CellRange referencing the category (should be
 * the same for all Series)
 * @param CelLRange  seriesRange - The CellRange referencing the data points for one
 * bar, line or wedge in the chart
 * @param CellRange  bubbleRange -The CellRange referencing bubble sizes for this
 * series, or null if chart is not of type BUBBLE
 * @return ChartSeriesHandle referencing the newly added series
 * @see ChartHandle.addSeriesRange
 */
     fun addSeriesRange(legendCell:CellHandle, categoryRange:CellRange, seriesRange:CellRange, bubbleRange:CellRange):ChartSeriesHandle {
return this
.addSeriesRange(legendCell, categoryRange, seriesRange, bubbleRange, 0) // 0=default
 // chart
    }

/**
 * Adds a new Series to the chart via CellHandles and CellRange Objects. Each
 * bar, line or wedge in a chart represents a Series. <br></br>
 * This method can update the default chart (nChart==0) or overlay charts
 * (nChart 1-9)
 *
 * @param CellHandle legendCell - references the legend cell for this series
 * @param CellRange  categoryRange - The CellRange referencing the category (should be
 * the same for all Series)
 * @param CelLRange  seriesRange - The CellRange referencing the data points for one
 * bar, line or wedge in the chart
 * @param CellRange  bubbleRange -The CellRange referencing bubble sizes for this
 * series, or null if chart is not of type BUBBLE
 * @param nChart     number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @return ChartSeriesHandle referencing the newly added series
 * @see ChartHandle.addSeriesRange
 */
     fun addSeriesRange(legendCell:CellHandle, categoryRange:CellRange, seriesRange:CellRange, bubbleRange:CellRange?, nChart:Int):ChartSeriesHandle {
var s:Series? = null
if (bubbleRange == null)
s = mychart!!.addSeries(seriesRange.toString(), categoryRange
.toString(), "", (legendCell.workSheetName + "!"
+ legendCell.cellAddress), legendCell
.stringVal, nChart)
else
s = mychart!!.addSeries(seriesRange.toString(), categoryRange
.toString(), bubbleRange!!
.toString(), (legendCell.workSheetName + "!"
+ legendCell.cellAddress), legendCell
.stringVal, nChart)
setDimensionsRecord() // 20080417 KSC: update chart DIMENSIONS record
 // upon update of series
        return ChartSeriesHandle(s, workBookHandle)
}

/**
 * remove the Series (bar, line or wedge) at the desired index
 *
 * @param int    index - series index (valid values: 0 to
 * getAllChartSeriesHandles().length-1)
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @see getAllChartSeriesHandles
 */
    @JvmOverloads  fun removeSeries(index:Int, nChart:Int = -1) {
val seriesperchart = mychart!!.getAllSeries(nChart)
val seriestodelete = seriesperchart.get(index) as Series // series to
 // delete
        mychart!!.removeSeries(seriestodelete)
setDimensionsRecord()
}

/**
 * updates (replaces) every Chart Series (bar, line or wedge on the Chart) with
 * the data from the array of values, legends, bubble sizes (optional) and
 * category range. <br></br>
 * NOTE: all arrays must be the same size (the exception is the bubleSizeRanges
 * array, which may be null)
 *
 *
 * NOTE: String arrays come in reverse order from plugins, so this method adds
 * series LIFO i.e. reversed
 *
 * @param String[] valueRanges - Array of Cell Ranges representing the Values or Data
 * points for each series (bar, line or wedge) on the Chart
 * @param String[] legendCells - Array of Cell Addresses representing the legends for
 * each Series
 * @param String[] bubbleSizeRanges - Array of Cell ranges representing the bubble
 * sizes for the Chart, or null if chart is not of type BUBBLE
 * @param String   categoryRange - The Cell Range representing the categories (X
 * Axis) for the entire Chart
 */
     fun addAllSeries(valueRanges:Array<String>, legendCells:Array<String>, bubbleSizeRanges:Array<String>, categoryRange:String) {
addAllSeries(valueRanges, legendCells, bubbleSizeRanges, categoryRange, 0) // do
 // for
        // default
        // chart
    }

/**
 * updates (replaces) every Chart Series (bar, line or wedge on the Chart) with
 * the data from the array of values, legends, bubble sizes (optional) and
 * category range For the desired chart (0=default 1-9=overlay charts) <br></br>
 * NOTE: all arrays must be the same size (the exception is the bubleSizeRanges
 * array, which may be null)
 *
 *
 * NOTE: String arrays come in reverse order from plugins, so this method adds
 * series LIFO i.e. reversed
 *
 * @param String[] valueRanges - Array of Cell Ranges representing the Values or Data
 * points for each series (bar, line or wedge) on the Chart
 * @param String[] legendCells - Array of Cell Addresses representing the legends for
 * each Series
 * @param String[] bubbleSizeRanges - Array of Cell ranges representing the bubble
 * sizes for the Chart, or null if chart is not of type BUBBLE
 * @param String   categoryRange - The Cell Range representing the categories (X
 * Axis) for the entire Chart
 * @param nChart   number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 */
    private fun addAllSeries(valueRanges:Array<String>, legendCells:Array<String>, bubbleSizeRanges:Array<String>?, categoryRange:String, nChart:Int) {
 // first, remove all existing series
        val v = mychart!!.allSeries
for (i in v.indices)
{
mychart!!.removeSeries(v.get(i) as Series)
}
try
{
val chartMetrics = mychart!!.getMetrics(workBookHandle) // build
 // or
            // retrieve
            // Chart
            // Metrics
            // -->
            // dimensions
            // +
            // series
            // data
            // ...
            this.mychart!!.legend!!
.resetPos(chartMetrics.get("y"), chartMetrics
.get("h"), chartMetrics
.get("canvash"), legendCells.size)
}
catch (e:Exception) {}

setDimensionsRecord()
 // now add series
        val hasBubbles = (((bubbleSizeRanges != null && bubbleSizeRanges!!.size == valueRanges.size)))
for (i in valueRanges.indices.reversed())
{
try
{
if (!hasBubbles)
 // usual case
                    this.addSeriesRange(legendCells[i], categoryRange, valueRanges[i], null, nChart)
else
this.addSeriesRange(legendCells[i], categoryRange, valueRanges[i], bubbleSizeRanges!![i], nChart)
}
catch (e:Exception) {
Logger.logErr("Error adding series: " + e.toString())
}

}
}

/**
 * Appends a series one row below the last series in the chart for the desired
 * chart
 *
 *
 * This can be utilized when programmatically adding rows of data that should be
 * reflected in the chart. <br></br>
 * Legend cell will be incremented by one row if a reference. Category range
 * will stay the same.
 *
 * <br></br>
 * In order for this method to work properly the chart must have row-based
 * series. If your chart utilizes column-based series, then you need to append a
 * category.
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @return ChartSeriesHandle representing newly added series
 * @see ChartHandle.appendRowCategoryToChart
 */
    @JvmOverloads  fun appendRowSeriesToChart(nChart:Int = 0):ChartSeriesHandle? {
val handles = this.getAllChartSeriesHandles(nChart)
val theHandle = handles[handles.size - 1]
var legendRef = theHandle.seriesLegendReference
if (legendRef != null && legendRef != "")
{
val sheetnm = legendRef!!.substring(0, legendRef!!.indexOf("!"))
legendRef = legendRef!!
.substring(legendRef!!.indexOf("!") + 1)
val rc = ExcelTools.getRowColFromString(legendRef!!)
rc[0] = rc[0] + 1
legendRef = sheetnm + "!" + ExcelTools.formatLocation(rc)
}
else if (legendRef == null)
{
legendRef = theHandle.seriesLegend
}
else
{
legendRef = ""
}
val categoryRange = theHandle.categoryRange
var seriesRange = theHandle.seriesRange
val sheetnm = seriesRange.substring(0, seriesRange.indexOf("!"))
seriesRange = seriesRange
.substring(seriesRange.indexOf("!") + 1)
val rc = ExcelTools.getRangeRowCol(seriesRange)
 // fiddle it, since exceltools doesn't translate back/forth
        val newRc = IntArray(4)
newRc[0] = rc[1]
newRc[1] = rc[0] + 1
newRc[2] = rc[3]
newRc[3] = rc[2] + 1
seriesRange = sheetnm + "!" + ExcelTools.formatRange(newRc)
try
{
return this
.addSeriesRange(legendRef, "", categoryRange, seriesRange, "", nChart)
}
catch (e:CellNotFoundException) {
Logger.logErr(("ChartHandle.appendRowSeriesToChart: Unable to append series to chart: " + e))
}

return null
}

/**
 * Append a row of categories to the bottom of the chart. <br></br>
 * Expands all Series to include the new bottom row. <br></br>
 * To be utilized when expanding a chart to encompass more data that has a
 * col-based series.
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @see ChartHandle.appendRowSeriesToChart
 */
    @JvmOverloads  fun appendRowCategoryToChart(nChart:Int = 0) {
val handles = this.getAllChartSeriesHandles(nChart)

for (i in handles.indices)
{
val theHandle = handles[i]
 // update the series
            var seriesRange = theHandle.seriesRange
var s = ExcelTools.stripSheetNameFromRange(seriesRange)
var sheetnm = s[0]
seriesRange = s[1]
 // String sheetnm = seriesRange.substring(0,
            // seriesRange.indexOf("!"));
            // seriesRange = seriesRange.substring(
            // seriesRange.indexOf("!")+1,
            // seriesRange.length());
            // Strip 2nd sheet ref, if any 20080213 KSC
            // int n= seriesRange.indexOf('!');
            // int m= seriesRange.indexOf(':');
            // seriesRange= seriesRange.substring(0, m+1) +
            // seriesRange.substring(n+1);

            var rc = ExcelTools.getRangeRowCol(seriesRange)
var newRc = IntArray(4)
newRc[0] = rc[1]
newRc[1] = rc[0]
newRc[2] = rc[3]
newRc[3] = rc[2] + 1
seriesRange = sheetnm + "!" + ExcelTools.formatRange(newRc)
theHandle.seriesRange = seriesRange
 // update the category
            seriesRange = theHandle.categoryRange
s = ExcelTools.stripSheetNameFromRange(seriesRange)
sheetnm = s[0]
seriesRange = s[1]
 /*
             * sheetnm = seriesRange.substring(0,
             * seriesRange.indexOf("!")); seriesRange =
             * seriesRange.substring( seriesRange.indexOf("!")+1,
             * seriesRange.length()); //
             * Strip 2nd sheet ref, if any 20080213 KSC n=
             * seriesRange.indexOf('!'); m=
             * seriesRange.indexOf(':'); seriesRange=
             * seriesRange.substring(0, m+1) +
             * seriesRange.substring(n+1);
             */
            rc = ExcelTools.getRangeRowCol(seriesRange)
newRc = IntArray(4)
newRc[0] = rc[1]
newRc[1] = rc[0]
newRc[2] = rc[3]
newRc[3] = rc[2] + 1
seriesRange = sheetnm + "!" + ExcelTools.formatRange(newRc)
theHandle.categoryRange = seriesRange
}
}

private fun t(n:Int):String {
val tabs = "\t\t\t\t\t\t\t\t\t\t\t\t\t"
return (tabs.substring(0, n))
}

/**
 * given a XmlPullParser positioned at the chart element, parse all chart
 * elements to create desired chart
 *
 * @param sht WorkSheetHandle
 * @param xpp XmlPullParser
 */
     fun parseXML(sht:WorkSheetHandle, xpp:XmlPullParser, maps:HashMap<*, *>) {
try
{
var eventType = xpp.getEventType()
while (eventType != XmlPullParser.END_DOCUMENT)
{
if (eventType == XmlPullParser.START_TAG)
{
val tnm = xpp.getName()
if (tnm == "ChartFonts")
{
for (x in 0 until xpp.getAttributeCount())
{
this.setChartFont(xpp.getAttributeName(x), xpp
.getAttributeValue(x))
}
}
else if (tnm == "ChartFontRec")
{
var fName = ""
var fontId = 0
var fSize = 0
var fWeight = 0
var fColor = 0
var fUnderline = 0
var bIsBold = false
for (x in 0 until xpp.getAttributeCount())
{
val attr = xpp.getAttributeName(x)
val `val` = xpp.getAttributeValue(x)
if (attr == "name")
{
fName = `val`
}
else if (attr == "id")
{
fontId = Integer.parseInt(`val`)
}
else if (attr == "size")
{
fSize = Font.PointsToFontHeight(java.lang.Double
.parseDouble(`val`))
}
else if (attr == "color")
{
fColor = FormatHandle
.HexStringToColorInt(`val`, FormatHandle.colorFONT)
if (fColor == 0)
fColor = 32767 // necessary?
}
else if (attr == "weight")
{
fWeight = Integer.parseInt(`val`)
}
else if (attr == "bold")
{
bIsBold = true
}
else if (attr == "underline")
{
fUnderline = Integer.parseInt(`val`)
}
}
while (this.workBook!!.numFonts < fontId - 1)
{
this.workBook!!.insertFont(Font("Arial",
FormatConstants.PLAIN, 10))
}
if (this.workBook!!.numFonts < fontId)
{
val f = Font(fName, fWeight, fSize)
f.color = fColor
f.bold = bIsBold
f.setUnderlineStyle(fUnderline.toByte())
this.workBook!!.insertFont(f)
}
else
{ // TODO: this will screw up linked fonts,
 // perhaps, so what to do?
                            val f = this.workBook!!.getFont(fontId)
f.fontWeight = fWeight
f.setFontName(fName)
f.fontHeight = fSize
f.color = fColor
f.bold = bIsBold
f.setUnderlineStyle(fUnderline.toByte())
}
}
else if (tnm == "FormatChartArea")
{ // TODO:
 // something!
                        // ChartBorder
                        // ChartProperties
                    }
else if (tnm == "Series")
{ // series -->
 // Legend Range Category shape typex typey
                        var legend = ""
var series = ""
var category = ""
var bubble = ""
var dataTypeX = ""
var dataTypeY = ""
var shape = ""
for (x in 0 until xpp.getAttributeCount())
{
if (xpp.getAttributeName(x)
.equals("Legend", ignoreCase = true))
legend = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("Range", ignoreCase = true))
series = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("Category", ignoreCase = true))
category = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("Bubbles", ignoreCase = true))
bubble = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("TypeX", ignoreCase = true))
dataTypeX = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("TypeY", ignoreCase = true))
dataTypeY = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("Shape", ignoreCase = true))
shape = xpp.getAttributeValue(x)
}
 // 20070709 KSC: can't add until all cells are added
                        // ch.addSeriesRange(legend, category, series);
                        val s = arrayOf<String>(legend, series, category, bubble, dataTypeX, dataTypeY, shape)
val map = maps
.get("chartseries") as HashMap<Array<String>, ChartHandle>
map.put(s, this)
}
else if (tnm == "ChartOptions")
{ // handle
 // chart-type-specific
                        // options such
                        // as legend
                        // options
                        // 20070716 KSC: handle chart-type-specific options in a
                        // very generic way ...
                        for (x in 0 until xpp.getAttributeCount())
{
this.setChartOption(xpp.getAttributeName(x), xpp
.getAttributeValue(x))
}
}
else if (tnm == "ThreeD")
{ // handle three-d options
 // handle threeD record options in a very generic way
                        // ...
                        this.make3D() // default chart - TODO; if mutliple
 // charts, handle
                        for (x in 0 until xpp.getAttributeCount())
{ // now
 // add
                            // threed
                            // rec
                            // options
                            this.setChartOption(xpp.getAttributeName(x), xpp
.getAttributeValue(x))
}
}
else if (tnm.endsWith("Axis"))
{ // handle axis specs
 // (Label + options ...)
                        // 20070720 KSC: handle Axis record options ...
                        var type = 0
val axis = xpp.getName()
if (axis.equals("XAxis", ignoreCase = true))
type = XAXIS
else if (axis.equals("YAxis", ignoreCase = true))
type = YAXIS
else if (axis.equals("ZAxis", ignoreCase = true))
type = ZAXIS
if (xpp.getAttributeCount() > 0)
{ // then has axis
 // options
                            for (x in 0 until xpp.getAttributeCount())
{
this.setAxisOption(type, xpp
.getAttributeName(x), xpp
.getAttributeValue(x))
}
}
else
{ // no axis options means no axis present;
 // remove
                            this.removeAxis(type)
}
}
else if (tnm == "Series")
{ // handle series data
 // Legend Range Category
                        var legend = ""
var series = ""
var category = ""
var bubble = ""
var dataTypeX = ""
var dataTypeY = ""
var shape = ""
for (x in 0 until xpp.getAttributeCount())
{
if (xpp.getAttributeName(x)
.equals("Legend", ignoreCase = true))
legend = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("Range", ignoreCase = true))
series = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("Category", ignoreCase = true))
category = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("Bubbles", ignoreCase = true))
bubble = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("TypeX", ignoreCase = true))
dataTypeX = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("TypeY", ignoreCase = true))
dataTypeY = xpp.getAttributeValue(x)
else if (xpp.getAttributeName(x)
.equals("Shape", ignoreCase = true))
shape = xpp.getAttributeValue(x)
}
 // 20070709 KSC: can't add until all cells are added
                        // ch.addSeriesRange(legend, category, series);
                        val s = arrayOf<String>(legend, series, category, bubble, dataTypeX, dataTypeY, shape)
val map = maps
.get("chartseries") as HashMap<Array<String>, ChartHandle>
map.put(s, this)
}
}
else if (eventType == XmlPullParser.END_TAG)
{
if (xpp.getName() == "Chart")
break
}
eventType = xpp.next()
}
}
catch (e:Exception) {
Logger.logWarn(("ChartHandle.parseXML <" + xpp.getName() + ">: "
+ e.toString()))
 // TODO: propogate Exception???
        }

}

/********************************************************************
 * OOXML Generation Methods
 */
    /**
 * Generates OOXML (chartML) for this chart object.
 *
 * <br></br>
 * NOTE: necessary root chartSpace element + namespaces are not set here
 *
 * @param int rId -reference ID for this chart
 * @return String representing the OOXML describing this Chart
 */
     fun getOOXML(rId:Int):String {
 // TODO: finish 3d options- floor, sideWall, backWall
        // TODO: finish axis options
        // TODO: printSettings

        // generate OOXML (chartML)
        val cooxml = StringBuffer()
try
{
mychart!!.chartSeries.resetSeriesNumber() // reset series idx
 // retrieve pertinent chart data
            // axes id's TODO: HANDLE MULTIPLE AXES per chart ...
            val catAxisId = Integer
.toString((Math.random() * 1000000).toInt())
val valAxisId = Integer
.toString((Math.random() * 1000000).toInt())
val serAxisId = Integer
.toString((Math.random() * 1000000).toInt())
val thischart:OOXMLChart
if ((mychart is OOXMLChart))
thischart = mychart as OOXMLChart?
else
{ // XLS->XLSX
thischart = OOXMLChart(mychart!!, workBookHandle)
mychart = thischart
thischart.chartSeries.setParentChart(thischart)
}
thischart.wbh = this.workBookHandle

cooxml.append(thischart.getOOXML(catAxisId, valAxisId, serAxisId))

 // TODO: <printSettings>
            val chartEmbeds = thischart.chartEmbeds
if (chartEmbeds != null)
{
var j = 0
for (i in chartEmbeds!!.indices)
{
if ((chartEmbeds!!.get(i) as Array<String>)[0] == "userShape")
{
j++
cooxml.append("<c:userShapes r:id=\"rId" + j + "\"/>")
}
}
}
}
catch (e:Exception) {
Logger.logErr(("ChartHandle.getOOXML: error generating OOXML.  Chart not created: " + e.toString()))
}

return cooxml.toString()
}

/**
 * generates the OOXML specific for DrawingML, specifying offsets and
 * identifying the chart object. <br></br>
 * this Drawing ML (OOXML) is distinct from Chart ML (OOXML) which actually
 * defines the chart object including series, categories and axes <br></br>
 * This is an internal method that is not useful to the end user.
 *
 * @param int id - the reference id for this chart
 * @return String OOXML
 */
     fun getChartDrawingOOXML(id:Int):String {
val t = TwoCellAnchor(
(mychart as OOXMLChart).editMovement)
t.setAsChart(id, OOXMLAdapter.stripNonAscii(this.ooxmlName)
.toString(), TwoCellAnchor.convertBoundsFromBIFF8(this
.sheet, mychart!!.bounds!!)) // adjust BIFF8
 // bounds to
        // OOXML units
        return t.ooxml
}

/********************************************************************************
 * Parsing OOXML Methods
 */
    /**
 * defines this chart object based on a Chart ML (OOXML) input Stream (root
 * element=c:chartSpace) <br></br>
 * This is an internal method that is not useful to the end user.
 *
 * @param inputStream ii - representing chart OOXML
 */
     fun parseOOXML(ii:java.io.InputStream) {
 // overlay in title, after layout
        // varyColors val= "0" -- after grouping and before ser

        // series colors by theme:
        /*
         * accent1 - 6
         */
        /**
 * chartSpace: chart (Chart) §5.7.2.27 clrMapOvr (Color Map Override) §5.7.2.30
 * date1904 (1904 Date System) §5.7.2.38 externalData (External Data
 * Relationship) §5.7.2.63 extLst (Chart Extensibility) §5.7.2.64 lang (Editing
 * Language) §5.7.2.87 pivotSource (Pivot Source) §5.7.2.145 printSettings
 * (Print Settings) §5.7.2.149 protection (Protection) §5.7.2.150 roundedCorners
 * (Rounded Corners) §5.7.2.160 spPr (Shape Properties) §5.7.2.198 style (Style)
 * §5.7.2.203 txPr (Text Properties) §5.7.2.217 userShapes (Reference to Chart
 * Drawing Part) §5.7.2.222
 */
        try
{
val thischart = mychart as OOXMLChart?
var drawingOrder = 0 // drawing order of the chart (0=default, 1-9
 // for multiple charts in 1)
            var hasPivotTableSource = false

 // remove any undesired formatting from default chart:
            this.title = "" // clear any previously set
mychart!!.axes!!.setPlotAreaBgColor(FormatConstants.COLOR_WHITE)
mychart!!.axes!!.setPlotAreaBorder(-1, -1) // remove plot area
 // border

            val lastTag = java.util.Stack<String>() // keep
 // track
            // of
            // element
            // hierarchy

            val factory = XmlPullParserFactory.newInstance()
factory.setNamespaceAware(true)
val xpp = factory.newPullParser()

xpp.setInput(ii, null) // using XML 1.0 specification
var eventType = xpp.getEventType()
while (eventType != XmlPullParser.END_DOCUMENT)
{
if (eventType == XmlPullParser.START_TAG)
{
val tnm = xpp.getName() // main entry= chartSpace,
 // children: lang, chart
                    lastTag.push(tnm) // keep track of element hierarchy
if (tnm == "chart")
{ // beginning of DrawingML for a
 // single image or chart;
                        // children: title,
                        // plotArea, legend,
                        /**
 * <sequence> 5
 * <element name="title" type="CT_Title" minOccurs="0" maxOccurs="1"></element> 6
 * <element name="autoTitleDeleted" type="CT_Boolean" minOccurs="0" maxOccurs=
                          "1"></element> 7
 * <element name="pivotFmts" type="CT_PivotFmts" minOccurs="0" maxOccurs="1"></element> 8
 * <element name="view3D" type="CT_View3D" minOccurs="0" maxOccurs="1"></element> 9
 * <element name="floor" type="CT_Surface" minOccurs="0" maxOccurs="1"></element> 10
 * <element name="sideWall" type="CT_Surface" minOccurs="0" maxOccurs="1"></element> 11
 * <element name="backWall" type="CT_Surface" minOccurs="0" maxOccurs="1"></element> 12
 * <element name="plotArea" type="CT_PlotArea" minOccurs="1" maxOccurs="1"></element>
 * plotArea contains layout, <ChartType>, serAx, valAx, catAx, dateAx, spPr,
 * dTable 13
 * <element name="legend" type="CT_Legend" minOccurs="0" maxOccurs="1"></element> 14
 * <element name="plotVisOnly" type="CT_Boolean" minOccurs="0" maxOccurs="1"></element>
 * 15
 * <element name="dispBlanksAs" type="CT_DispBlanksAs" minOccurs="0" maxOccurs=
                          "1"></element> 16
 * <element name="showDLblsOverMax" type="CT_Boolean" minOccurs="0" maxOccurs=
                          "1"></element> 17
 * <element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"></element>
 * 18 </ChartType></sequence>
 */
                    }
else if (tnm == "lang")
{
thischart!!.lang = xpp.getAttributeValue(0)
}
else if (tnm == "roundedCorners")
{
thischart!!.roundedCorners = xpp.getAttributeValue(0) == "1"
}
else if (tnm == "pivotSource")
{ // has a pivot table
hasPivotTableSource = true
}
else if (tnm == "view3D")
{
ThreeD.parseOOXML(xpp, lastTag, thischart!!)
}
else if (tnm == "layout")
{
thischart!!.plotAreaLayout = Layout
.parseOOXML(xpp, lastTag).cloneElement() as Layout
}
else if (tnm == "legend")
{
thischart!!.showLegend(true, false)
thischart!!.ooxmlLegend = io.starter.formats.OOXML.Legend
.parseOOXML(xpp, lastTag, this.workBookHandle)
.cloneElement() as io.starter.formats.OOXML.Legend
thischart!!.ooxmlLegend!!
.fill2003Legend(thischart!!.legend)
 // Parse actual CHART TYPE element (barChart, pieChart,
                        // etc.)
                    }
else if ((tnm == OOXMLConstants.twoDchartTypes[ChartHandle.BARCHART]
|| tnm == OOXMLConstants.twoDchartTypes[ChartHandle.LINECHART]
|| tnm == OOXMLConstants.twoDchartTypes[ChartHandle.PIECHART]
|| tnm == OOXMLConstants.twoDchartTypes[ChartHandle.AREACHART]
|| tnm == OOXMLConstants.twoDchartTypes[ChartHandle.SCATTERCHART]
|| tnm == OOXMLConstants.twoDchartTypes[ChartHandle.RADARCHART]
|| tnm == OOXMLConstants.twoDchartTypes[ChartHandle.SURFACECHART]
|| tnm == OOXMLConstants.twoDchartTypes[ChartHandle.DOUGHNUTCHART]
|| tnm == OOXMLConstants.twoDchartTypes[ChartHandle.BUBBLECHART]
|| tnm == OOXMLConstants.twoDchartTypes[ChartHandle.OFPIECHART]
|| tnm == OOXMLConstants.twoDchartTypes[ChartHandle.STOCKCHART]
|| tnm == OOXMLConstants.threeDchartTypes[ChartHandle.BARCHART]
|| tnm == OOXMLConstants.threeDchartTypes[ChartHandle.LINECHART]
|| tnm == OOXMLConstants.threeDchartTypes[ChartHandle.PIECHART]
|| tnm == OOXMLConstants.threeDchartTypes[ChartHandle.AREACHART]
|| tnm == OOXMLConstants.threeDchartTypes[ChartHandle.SCATTERCHART]
|| tnm == OOXMLConstants.threeDchartTypes[ChartHandle.SURFACECHART]))
{ // specific
 // chart
                        // type-
                        ChartType
.parseOOXML(xpp, this.workBookHandle, mychart, drawingOrder++)
lastTag.pop()
}
else if (tnm == "title")
{
thischart!!.setOOXMLTitle(Title
.parseOOXML(xpp, lastTag, this.workBookHandle)
.cloneElement() as Title, this.workBookHandle)
this.title = thischart!!.ooxmlTitle!!.title
}
else if (tnm == "spPr")
{ // shape properties -- can
 // be for plot area or
                        // chart space
                        val parent = lastTag.get(lastTag.size - 2)
if (parent == "plotArea")
{
thischart!!.setSpPr(0, SpPr
.parseOOXML(xpp, lastTag, this.workBookHandle)
.cloneElement() as SpPr)
}
else if (parent == "chartSpace")
{
thischart!!.setSpPr(1, SpPr
.parseOOXML(xpp, lastTag, this.workBookHandle)
.cloneElement() as SpPr)
}
}
else if (tnm == "txPr")
{ // text formatting
thischart!!.txPr = TxPr
.parseOOXML(xpp, lastTag, this.workBookHandle)
.cloneElement() as TxPr
}
else if (tnm == "catAx")
{ // child of plotArea
mychart!!.axes!!
.parseOOXML(XAXIS, xpp, tnm, lastTag, this.workBookHandle)
}
else if (tnm == "valAx")
{ // child of plotArea
if (mychart!!.axes!!.hasAxis(ChartConstants.XAXIS))
 // usual,
                            // have
                            // a
                            // catAx
                            // then
                            // a
                            // valAx
                            mychart!!.axes!!
.parseOOXML(ChartConstants.YAXIS, xpp, tnm, lastTag, this.workBookHandle)
else if (mychart!!.axes!!
.hasAxis(ChartConstants.YAXIS))
 // for bubble
                            // charts, has
                            // two valAxes
                            // and no
                            // catAx
                            mychart!!.axes!!
.parseOOXML(ChartConstants.XVALAXIS, xpp, tnm, lastTag, this.workBookHandle)
else
 // 2nd val axis is Y axis
                            mychart!!.axes!!
.parseOOXML(ChartConstants.YAXIS, xpp, tnm, lastTag, this.workBookHandle)
}
else if (tnm == "serAx")
{ // series axis - 3d charts
mychart!!.axes!!
.parseOOXML(ZAXIS, xpp, tnm, lastTag, this.workBookHandle)
}
else if (tnm == "dateAx")
{ // TODO: not finished:
 // figure out!
                        // ??
                    }

}
else if (eventType == XmlPullParser.END_TAG)
{
lastTag.pop()
val endTag = xpp.getName()
if (endTag == "chartSpace")
{
setDimensionsRecord()
break // done processing
}
}
if (xpp.getEventType() != XmlPullParser.END_DOCUMENT)
eventType = xpp.next()
else
eventType = XmlPullParser.END_DOCUMENT
}
}
catch (e:Exception) {
Logger.logErr("ChartHandle.parseChartOOXML: " + e.toString())
}

}

/**
 * Specifies how to resize or move this Chart upon edit <br></br>
 * This is an internal method that is not useful to the end user. <br></br>
 * Excel 7/OOXML specific
 *
 * @param editMovement String OOXML-specific edit movement setting
 */
     fun setEditMovement(editMovement:String) {
(mychart as OOXMLChart).editMovement = editMovement
}

/**
 * sets external information linked to or "embedded" in this OOXML chart; can be
 * a chart user shape, an image ... <br></br>
 * NOTE: a userShape is a drawingml file name which defines the userShape (if
 * any) <br></br>
 * a userShape is a drawing or shape ontop of a chart
 *
 * @param String[] embedType, filename e.g. {"userShape", "userShape file name"}
 */
     fun addChartEmbed(ce:Array<String>) {
(mychart as OOXMLChart).addChartEmbed(ce)
}

/**
 * set the chart DIMENSIONS record based on the series ranges in the chart
 * APPEARS THAT for charts, the DIMENSIONS record merely notes the range of
 * values: 0, #points in series, 0, #series
 */
    protected fun setDimensionsRecord() {
val series = this.allChartSeriesHandles
val nSeries = series.size
var nPoints = 0
for (i in series.indices)
{
try
{
val coords = ExcelTools
.getRangeCoords(series[i].seriesRange)
if (coords[3] > coords[1])
nPoints = Math.max(nPoints, coords[3] - coords[1] + 1) // c1-c0
else
nPoints = Math.max(nPoints, coords[2] - coords[0] + 1) // r1-r0
}
catch (e:Exception) {}

}
mychart!!.setDimensionsRecord(0, nPoints, 0, nSeries)
}

/**
 * Method for setting Chart-Type-specific options in a generic fashion e.g.
 * charthandle.setChartOption("Stacked", "true");
 *
 *
 * Note: since most Chart Type Options are interdependent, there are several
 * makeXX methods that set the desired group of options e.g. makeStacked(); use
 * setChartOption with care
 *
 *
 * Note that not all Chart Types will have every option available
 *
 *
 * Possible Options:
 *
 *
 * "Stacked" - true or false - set Chart Series to be Stacked <br></br>
 * "Cluster" - true or false - set Clustered for Column and Bar Chart Types <br></br>
 * "PercentageDisplay" - true or false - Each Category is broken down as a
 * percentge <br></br>
 * "Percentage" - Distance of pie slice from center of pie as % for Pie Charts
 * (0 for all others) <br></br>
 * "donutSize" - Donut size for Donut Charts Only <br></br>
 * "Overlap" - Space between bars (default= 0%) <br></br>
 * "Gap" - Space between categories (%) (default=50%) <br></br>
 * "SmoothedLine" - true or false - the Line series has a smoothed line <br></br>
 * "AnRot" - Rotation Angle (0 to 360 degrees), usually 0 for pie, 20 for others
 * (3D option) <br></br>
 * "AnElev" - Elevation Angle (-90 to 90 degrees) (15 is default) (3D option)
 * <br></br>
 * "ThreeDScaling" - true or false - 3d effect <br></br>
 * "TwoDWalls" - true if 2D walls (3D option) <br></br>
 * "PcDist" - Distance from eye to chart (0 to 100) (30 is default) (3D option)
 * <br></br>
 * "ThreeDBubbles" - true or false - Draw bubbles with a 3d effect <br></br>
 * "ShowLdrLines" - true or false - Show Pie and Donut charts Leader Lines <br></br>
 * "MarkerFormat" - "0" thru "9" for various marker options @see
 * ChartHandle.setMarkerFormat <br></br>
 * "ShowLabel" - true or false - show Series/Data Label <br></br>
 * "ShowCatLabel" - true or false - show Category Label <br></br>
 * "ShowLabelPct" - true or false - show percentage labels for Pie charts <br></br>
 * "ShowBubbleSizes" - true or false - show bubble sizes for Bubble charts
 *
 *
 * NOTE: all values must be in String form
 *
 * @see ChartHandle.getXML
 */
     fun setChartOption(op:String, `val`:String) {
mychart!!.setChartOption(op, `val`)
}

/**
 * Method for setting Chart-Type-specific options in a generic fashion e.g.
 * charthandle.setChartOption("Stacked", "true");
 *
 * @param op
 * @param val
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 */
    private fun setChartOption(op:String, `val`:String, nChart:Int) {
mychart!!.setChartOption(op, `val`, nChart)
}

/**
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @return true if Chart has 3D effects, false otherwise
 */
     fun isThreeD(nChart:Int):Boolean {
return mychart!!.isThreeD(nChart)
}

/**
 * Make chart 3D if not already
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 * @return ThreeD rec
 */
    internal fun initThreeD(nChart:Int):ThreeD {
return mychart!!.initThreeD(nChart, this.getChartType(nChart))
}

/**
 * @param int Axis - one of the Axis constants (XAXIS, YAXIS or ZAXIS)
 * @return String XML representation of the desired Axis
 */
    private fun getAxisOptionsXML(Axis:Int):String {
return mychart!!.axes!!.getAxisOptionsXML(Axis)
}

/**
 * Returns the Axis Label Placement or position as an int
 *
 *
 * One of: <br></br>
 * Axis.INVISIBLE - axis is hidden <br></br>
 * Axis.LOW - low end of plot area <br></br>
 * Axis.HIGH - high end of plot area <br></br>
 * Axis.NEXTTO- next to axis (default)
 *
 * @param int Axis - one of the Axis constants (XAXIS, YAXIS or ZAXIS)
 * @return int - one of the Axis Label placement constants above
 */
     fun getAxisPlacement(Axis:Int):Int {
return mychart!!.axes!!.getAxisPlacement(Axis)
}

/**
 * Sets the Axis labels position or placement to the desired value (these match
 * Excel placement options)
 *
 *
 * Possible options: <br></br>
 * Axis.INVISIBLE - hides the axis <br></br>
 * Axis.LOW - low end of plot area <br></br>
 * Axis.HIGH - high end of plot area <br></br>
 * Axis.NEXTTO- next to axis (default)
 *
 * @param int       Axis - one of the Axis constants (XAXIS, YAXIS or ZAXIS)
 * @param Placement - int one of the Axis placement constants listed above
 */
     fun setAxisPlacement(Axis:Int, Placement:Int) {
mychart!!.axes!!.setAxisPlacement(Axis, Placement)
mychart!!.setDirtyFlag(true)
}

 /*
     * returns the desired axis If bCreateIfNecessary, will
     * creates if doesn't exist
     *
     * @return Axis Object / private Axis getAxis(int axisType,
     * boolean
     * bCreateIfNecessary) { return mychart.getAxis(axisType,
     * bCreateIfNecessary); }
     */

    /**
 * removes the desired Axis from the Chart
 *
 * @param int axisType - one of the Axis constants (XAXIS, YAXIS or ZAXIS)
 */
     fun removeAxis(axisType:Int) {
mychart!!.axes!!.removeAxis(axisType)
mychart!!.setDirtyFlag(true)
}

/**
 * Set non-axis chart font id for title, default, etc <br></br>
 * For Internal Use Only
 *
 * @param String type - font type
 * @param String val - font id
 */
     fun setChartFont(type:String, `val`:String) {
mychart!!.setChartFont(type, `val`)
}

 // 20070802 KSC: Debug uility to write out chartrecs
    /*
     * For internal debugging use only
     *
     * public void writeChartRecs(String fName) {
     * mychart.writeChartRecs(fName); }
     *
     * /** set DataLabels option for this chart
     *
     * @param String type - see below
     *
     * @param boolean bShowLegendKey - true if show legend key,
     * false otherwise
     * <p>possible String type values: <br>Series <br>Category
     * <br>Value
     * <br>Percentage (Only for Pie Charts) <br>Bubble (Only for
     * Bubble Charts)
     * <br>X Value (Only for Bubble Charts) <br>Y Value (Only
     * for Bubble Charts)
     * <br>CandP
     *
     * <br><br>NOTE: not 100% implemented at this time
     */
     fun setDataLabel(/* [] */ type:String, bShowLegendKey:Boolean) {
 /*
         * for now, only 1 option is valid - multiple legend
         * settings e.g. Category and
         * Value are not figured out yet for (int i= 0; i <
         * type.length; i++)
         * mychart.setChartOption("DataLabel", type[i]);
         */
        if (!bShowLegendKey)
mychart!!.setChartOption("DataLabel", type)
else
mychart!!.setChartOption("DataLabelWithLegendKey", type)
}

/**
 * shows or removes the Data Table for this chart
 *
 * @param boolean bShow - true if show data table
 */
     fun showDataTable(bShow:Boolean) {
mychart!!.showDataTable(bShow)
}

/**
 * shows or hides the Chart legend key
 *
 * @param booean  bShow - true if show legend, false to hide
 * @param boolean vertical - true if show vertically, false for horizontal
 */
     fun showLegend(bShow:Boolean, vertical:Boolean) {
mychart!!.showLegend(bShow, vertical)
}

/**
 * returns true if Chart has a Data Legend Key showing
 *
 * @return true if Chart has a Data Legend Key showing
 */
     fun hasDataLegend():Boolean {
return mychart!!.hasDataLegend()
}

 fun removeLegend() {
mychart!!.removeLegend()
}

 // 20070905 KSC: Group chart options for ease of setting
    // almost all charts have these specific ChartTypes:

    /**
 * makes this Chart Stacked <br></br>
 * sets the group of options necessary to create a stacked chart <br></br>
 * For Chart Types: <br></br>
 * BAR, COL, LINE, AREA, PYRAMID, PYRAMIDBAR, CYLINDER, CYLINDERBAR, CONE,
 * CONEBAR
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 */
    @Deprecated("")
 fun makeStacked(nChart:Int) { // bar, col, line, area, pyramid col +
 // bar, cone col + bar, cylinder col
        // + bar
        val chartType = this.getChartType(nChart)
this.setChartOption("Stacked", "true", nChart)
when (chartType) {
ChartConstants.BARCHART, ChartConstants.COLCHART -> this.setChartOption("Overlap", "-100", nChart)
ChartConstants.CYLINDERCHART, ChartConstants.CYLINDERBARCHART, ChartConstants.CONECHART, ChartConstants.CONEBARCHART, ChartConstants.PYRAMIDCHART, ChartConstants.PYRAMIDBARCHART -> {
this.setChartOption("Overlap", "-100", nChart)
val td = this.initThreeD(nChart)
td.setChartOption("Cluster", "false")
}
ChartConstants.LINECHART -> this.setChartOption("Percentage", "0", nChart)
ChartConstants.AREACHART -> {
this.setChartOption("Overlap", "-100", nChart)
this.setChartOption("Percentage", "25", nChart)
this.setChartOption("SmoothedLine", "true", nChart)
}
}
}

/**
 * makes this Chart 100% Stacked <br></br>
 * sets the group of options necessary to create a 100% stacked chart <br></br>
 * For Chart Types: <br></br>
 * BAR, COL, LINE, PYRAMID, PYRAMIDBAR, CYLINDER, CYLINDERBAR, CONE, CONEBAR
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 */
    @Deprecated("")
 fun make100PercentStacked(nChart:Int) { // bar, col, line, pyramid
 // col + bar, cone col +
        // bar, cylinder col +
        // bar
        val chartType = this.getChartType(nChart)
this.setChartOption("Stacked", "true", nChart)
this.setChartOption("PercentageDisplay", "true", nChart)
when (chartType) {
ChartConstants.COLCHART // + pyramid
, ChartConstants.BARCHART -> this.setChartOption("Overlap", "-100", nChart)
ChartConstants.CYLINDERCHART, ChartConstants.CYLINDERBARCHART, ChartConstants.CONECHART, ChartConstants.CONEBARCHART, ChartConstants.PYRAMIDCHART, ChartConstants.PYRAMIDBARCHART -> {
this.setChartOption("Overlap", "-100", nChart)
val td = this.initThreeD(nChart)
td.setChartOption("Cluster", "false")
}
ChartConstants.LINECHART -> {}
}
}

/**
 * makes this Chart Stacked with a 3D Effect <br></br>
 * sets the group of options necessary to create a Stacked 3D chart <br></br>
 * For Chart Types: <br></br>
 * BAR, COL, AREA
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 */
     fun makeStacked3D(nChart:Int) { // bar, col, area
val chartType = this.getChartType(nChart)
this.setChartOption("Stacked", "true", nChart)
val td = this.initThreeD(nChart)
td.setChartOption("AnRot", "20")
td.setChartOption("ThreeDScaling", "true")
td.setChartOption("TwoDWalls", "true")
when (chartType) {
ChartConstants.COLCHART, ChartConstants.BARCHART -> this.setChartOption("Overlap", "-100", nChart)
ChartConstants.AREACHART -> {
this.setChartOption("Percentage", "25", nChart)
this.setChartOption("SmoothedLine", "true", nChart)
}
}
}

/**
 * makes this Chart 100% Stacked with a 3D Effect <br></br>
 * sets the group of options necessary to create a 100% Stacked 3D chart <br></br>
 * For Chart Types: <br></br>
 * BAR, COL, AREA
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 */
    @Deprecated("")
 fun make100PercentStacked3D(nChart:Int) { // bar, col, area
val chartType = this.getChartType(nChart)
this.setChartOption("Stacked", "true", nChart)
this.setChartOption("PercentageDisplay", "true", nChart)
when (chartType) {
ChartConstants.COLCHART // + pyramid
, ChartConstants.BARCHART -> {
this.setChartOption("Overlap", "-100", nChart)
val td = this.initThreeD(nChart)
td.setChartOption("AnRot", "20")
td.setChartOption("ThreeDScaling", "true")
td.setChartOption("TwoDWalls", "true")
}
ChartConstants.AREACHART -> {
this.setChartOption("Percentage", "25", nChart)
this.setChartOption("SmoothedLine", "true", nChart)
}
}
}

/**
 * makes the desired Chart hava a 3D effect <br></br>
 * where nChart 0= default, 1-9=multiple charts in one <br></br>
 * sets the group of options necessary to create a 3D chart <br></br>
 * For Chart Types: <br></br>
 * BARCHART, COLCHART, LINECHART, PIECHART, AREACHART, BUBBLECHART,
 * PYRAMIDCHART, CYLINDERCHART, CONECHART
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 */
    @Deprecated("use setChartType(chartType, nChart, is3d, isStacked,\n"+
"      is100%Stacked) instead")
@JvmOverloads  fun make3D(nChart:Int = 0) { // bar, col, line, pie, area, bubble,
 // pyramid, cone, cylinder
        val chartType = this.getChartType(nChart)
var td:ThreeD? = null
when (chartType) {
ChartConstants.COLCHART, ChartConstants.BARCHART -> {
td = this.initThreeD(nChart)
td!!.setChartOption("AnRot", "20")
td!!.setChartOption("ThreeDScaling", "true")
td!!.setChartOption("TwoDWalls", "true")
}
ChartConstants.CYLINDERCHART, ChartConstants.CONECHART, ChartConstants.PYRAMIDCHART -> {
td = this.initThreeD(nChart)
td!!.setChartOption("Cluster", "false")
}
ChartConstants.AREACHART -> {
this.setChartOption("Percentage", "25", nChart)
this.setChartOption("SmoothedLine", "true", nChart)
td = this.initThreeD(nChart)
td!!.setChartOption("AnRot", "20")
td!!.setChartOption("ThreeDScaling", "true")
td!!.setChartOption("TwoDWalls", "true")
td!!.setChartOption("Perspective", "true")
}
ChartConstants.PIECHART, ChartConstants.LINECHART -> this.initThreeD(nChart) // just create a threeD rec w/ no extra
ChartConstants.BUBBLECHART -> {
this.setChartOption("Percentage", "25", nChart)
this.setChartOption("SmoothedLine", "true", nChart)
this.setChartOption("ThreeDBubbles", "true", nChart)
td = this.initThreeD(nChart) // 20081228 KSC
}
}// options
}

 // more specialized option sets

    /**
 * makes this Chart Clusted with a 3D effect <br></br>
 * sets the group of options necessary to create a Clusted 3D chart <br></br>
 * For Chart Types: <br></br>
 * BAR, COL
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 */
    @Deprecated("")
 fun makeClustered3D(nChart:Int) { // only for Column and Bar (?)
val chartType = this.getChartType(nChart)
when (chartType) {
ChartConstants.BARCHART, ChartConstants.COLCHART -> {
val td = this.initThreeD(nChart)
td.setChartOption("AnRot", "20")
td.setChartOption("Cluster", "true")
td.setChartOption("ThreeDScaling", "true")
td.setChartOption("TwoDWalls", "true")
}
}
}

/**
 * makes this chart's wedges exploded i.e. separated <br></br>
 * For Chart Types: <br></br>
 * PIECHART, DOUGHNUTCHART
 *
 */
    @Deprecated("")
 fun makeExploded() { // pie, donut
val chartType = this.chartType
when (chartType) {
ChartConstants.DOUGHNUTCHART -> {
this.setChartOption("SmoothedLine", "true")
this.setChartOption("ShowLdrLines", "true")
this.setChartOption("Percentage", "25")
}
ChartConstants.PIECHART -> {
this.setChartOption("ShowLdrLines", "true")
this.setChartOption("Percentage", "25")
}
}// ShowLdrLines="true" Percentage="25"/>
 // exploded donut: ShowLdrLines="true" Donut="50"
 // Percentage="25"
 // SmoothedLine="true"/>
}

/**
 * makes this chart's wedges exploded 3D i.e. separated with a 3D effect <br></br>
 * For Chart Types: <br></br>
 * PIECHART, DOUGHNUTCHART
 *
 * @param nChart number and drawing order of the desired chart (default= 0 max=9
 * where 1-9 indicate an overlay chart)
 */
    @Deprecated("")
 fun makeExploded3D(nChart:Int) { // pie
 // ShowLdrLines="true" Percentage="25"
        // AnRot="236"
        val chartType = this.getChartType(nChart)
when (chartType) {
ChartConstants.DOUGHNUTCHART, ChartConstants.PIECHART -> {
this.setChartOption("ShowLdrLines", "true", nChart)
this.setChartOption("Percentage", "25", nChart)
val td = this.initThreeD(nChart)
td.setChartOption("AnRot", "236")
}
}
}

 /*
     * NOT IMPLEMENTED YET TODO: IMPLEMENT Make this Chart have
     * smoothed lines
     * (Scatter only)
     *
     * public void makeSmoothedLines() { // scatter
     * //Percentage="25"
     * SmoothedLine="true }
     *
     * public void makeWireFrame() { // surface // NO ColorFill,
     * only //
     * Percentage="25" SmoothedLine="true" // AnRot="20"
     * Perspective="true"
     * ThreeDScaling="true" TwoDWalls="true"/> // all else
     * should be default for
     * surface charts } public void makeContour() { // surface
     * -- for wireframe
     * surface, no ColorFill //ColorFill="true" Percentage="25"
     * SmoothedLine="true"/> // AnElev="90" pcDist="0"
     * Perspective="true"
     * ThreeDScaling="true" TwoDWalls="true" }
     */

    /**
 * set the marker format style for this chart <br></br>
 * one of: <br></br>
 * 0 = no marker <br></br>
 * 1 = square <br></br>
 * 2 = diamond <br></br>
 * 3 = triangle <br></br>
 * 4 = X <br></br>
 * 5 = star <br></br>
 * 6 = Dow-Jones <br></br>
 * 7 = standard deviation <br></br>
 * 8 = circle <br></br>
 * 9 = plus sign <br></br>
 * For Chart Types: <br></br>
 * LINE, SCATTER
 *
 * @param int imf - marker format constant from list above
 */
     fun setMarkerFormat(imf:Int) { // line, scatter ...
this.setChartOption("MarkerFormat", (imf).toString())
}

/**
 * utility to add a JSON object <br></br>
 * This is an internal method that is not useful to the end user.
 *
 * @param source
 * @param input
 */
    protected fun inputJSONObject(source:JSONObject?, input:JSONObject) {
if (source != null)
{
try
{
for (j in 0 until input.names().length())
{
source!!.put(input.names().getString(j), input
.get(input.names().getString(j)))
}
}
catch (e:JSONException) {
Logger.logErr("Error inputting JSON Object: " + e)
}

}
}

/**
 * /** Take current Chart object and return the SVG code necessary to define it,
 * scaled to the desired percentage e.g. 0.75= 75%
 *
 * @param scale double scale factor
 * @return String SVG
 */
     fun getSVG(scale:Double):String {
val chartMetrics = mychart!!.getMetrics(workBookHandle) // build
 // or
        // retrieve
        // Chart
        // Metrics
        // -->
        // dimensions
        // +
        // series
        // data
        // ...

        val svg = StringBuffer()
 // required header
        // svg.append("<?xml version=\"1.0\"
        // standalone=\"no\"?>\r\n"); // referneces
        // the DTD
        // svg.append("<!DOCTYPE svg PUBLIC \"-//W3C//DTD SVG
        // 1.1//EN\"
        // \"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd\">\r\n");

        // Define SVG Canvas:
        svg.append(("<svg width='" + (chartMetrics.get("canvasw") * scale)
+ "px' height='" + (chartMetrics.get("canvash") * scale)
+ "px' version=\"1.1\" xmlns=\"http://www.w3.org/2000/svg\"  xmlns:xlink=\"http://www.w3.org/1999/xlink\">\r\n"))
svg.append(("<g transform='scale(" + scale
+ ")'  onclick='null;' onmousedown='handleClick(evt);' style='width:100%; height:"
+ (chartMetrics.get("canvash") * scale) + "px'>")) // scale
 // chart
        // -default
        // scale=1
        // == 100%
        // JavaScript hooks
        svg.append(javaScript)

 // Data Legend Box -- do before drawing plot area as legends
        // box may change plot
        // area coordinates
        val legendSVG = getLegendSVG(chartMetrics) // but have to append it
 // after because should
        // overlap the
        // plotarea

        val bgclr = this.mychart!!.plotAreaBgColor
 // setup gradients
        svg.append("<defs>")
svg.append("<linearGradient id='bg_gradient' x1='0' y1='0' x2='0' y2='100%'>")
svg.append(("<stop offset='0' style='stop-color:" + bgclr
+ "; stop-opacity:1'/>"))
svg.append(("<stop offset='" + chartMetrics.get("w")
+ "' style='stop-color:white; stop-opacity:.5'/>"))
svg.append("</linearGradient>")
svg.append("</defs>")

 // PLOT AREA BG + RECT
        // rectangle around entire chart canvas
        if ((!(mychart is OOXMLChart) || !(mychart as OOXMLChart).roundedCorners))
svg.append(("<rect x='0' y='0' width='" + chartMetrics.get("canvasw")
+ "' height='" + chartMetrics.get("canvash")
+ "' style='fill-opacity:1;fill:white' stroke='#CCCCCC' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'/>\r\n"))
else
 // OOXML rounded corners
            svg.append(("<rect x='0' y='0' width='" + chartMetrics.get("canvasw")
+ "' height='" + chartMetrics.get("canvash")
+ "' rx='20' ry='20" + /* rounded corners */
"' style='fill-opacity:1;fill:white' stroke='#CCCCCC' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'/>\r\n"))

 // actual plot area
        svg.append(("<rect fill='" + bgclr
+ "'  style='fill-opacity:1;fill:url(#bg_gradient)' stroke='none' stroke-opacity='.5' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'"
+ " x='" + chartMetrics.get("x") + "' y='"
+ chartMetrics.get("y") + "' width='" + chartMetrics.get("w")
+ "' height='" + chartMetrics.get("h")
+ "' fill-rule='evenodd'/>\r\n"))

 // AXES, IF PRESENT - DO BEFORE ACTUAL SERIES DATA SO DRAWS
        // CORRECTLY + ADJUST
        // CHART DIMENSIONS
        svg.append(mychart!!.axes!!.getSVG(XAXIS, chartMetrics, mychart!!
.chartSeries.getCategories()))
svg.append(mychart!!.axes!!.getSVG(YAXIS, chartMetrics, mychart!!
.chartSeries.getCategories()))
 // TODO: Z Axis

        // After Axes and gridlines (if present),
        // ACTUAL bar/series/area/etc. svg generated from series and
        // scale data
        svg.append(this.mychart!!.chartObject.getSVG(chartMetrics, mychart!!
.axes!!.metrics, mychart!!.chartSeries))

svg.append(legendSVG) // append legend SVG obtained above

 // CHART TITLE

        if (mychart!!.titleTd != null)
svg.append(mychart!!.titleTd!!.getSVG(chartMetrics))

svg.append("</g>")
svg.append("</svg>")
 /*
         * //KSC: TESTING: REMOVE WHEN DONE if
         * (WorkBookFactory.PID==WorkBookFactory.E360) { // save svg
         * for testing
         * purposes try { java.io.File f= new java.io.File(
         * "c:/eclipse/workspace/testfiles/io.starter.OpenXLS/output/charts/FromSheetster.svg"
         * ); java.io.FileOutputStream fos= new
         * java.io.FileOutputStream(f);
         * fos.write(svg.toString().getBytes()); fos.flush();
         * fos.close(); } catch
         * (Exception e) {} }
         */
        return svg.toString()
}

private fun getLegendSVG(chartMetrics:HashMap<String, Double>):String? {
try
{
return mychart!!.legend!!.getSVG(chartMetrics, mychart!!
.chartObject, mychart!!.chartSeries)
}
catch (ne:NullPointerException) { // no legend??
return null
}

}

/**
 * debugging utility remove when done
 */
     fun WriteMainChartRecs(fName:String) {
mychart!!.chartObject.WriteMainChartRecs(fName)
}

 fun getAxis(type:Int):Axis? {
 // TODO Auto-generated method stub
        return null
}

companion object {
 // 20080114 KSC: delegate so visible
     val BARCHART = ChartConstants.BARCHART
 val COLCHART = ChartConstants.COLCHART
 val LINECHART = ChartConstants.LINECHART
 val PIECHART = ChartConstants.PIECHART
 val AREACHART = ChartConstants.AREACHART
 val SCATTERCHART = ChartConstants.SCATTERCHART
 val RADARCHART = ChartConstants.RADARCHART
 val SURFACECHART = ChartConstants.SURFACECHART
 val DOUGHNUTCHART = ChartConstants.DOUGHNUTCHART
 val BUBBLECHART = ChartConstants.BUBBLECHART
 val OFPIECHART = ChartConstants.OFPIECHART
 val PYRAMIDCHART = ChartConstants.PYRAMIDCHART
 val CYLINDERCHART = ChartConstants.CYLINDERCHART
 val CONECHART = ChartConstants.CONECHART
 val PYRAMIDBARCHART = ChartConstants.PYRAMIDBARCHART
 val CYLINDERBARCHART = ChartConstants.CYLINDERBARCHART
 val CONEBARCHART = ChartConstants.CONEBARCHART
 val RADARAREACHART = ChartConstants.RADARAREACHART
 val STOCKCHART = ChartConstants.STOCKCHART

 // legacy
     val BAR = BARCHART
 val COL = COLCHART
 val LINE = LINECHART
 val PIE = PIECHART
 val AREA = AREACHART
 val SCATTER = SCATTERCHART
 val RADAR = RADARCHART
 val SURFACE = SURFACECHART
 val DOUGHNUT = DOUGHNUTCHART
 val BUBBLE = BUBBLECHART
 val RADARAREA = RADARAREACHART
 val PYRAMID = PYRAMIDCHART
 val CYLINDER = CYLINDERCHART
 val CONE = CONECHART
 val PYRAMIDBAR = PYRAMIDBARCHART
 val CYLINDERBAR = CYLINDERBARCHART
 val CONEBAR = CONEBARCHART
 // axis types
     val XAXIS = ChartConstants.XAXIS
 val YAXIS = ChartConstants.YAXIS
 val ZAXIS = ChartConstants.ZAXIS
 val XVALAXIS = ChartConstants.XVALAXIS            // an
 // X
    // axis
    // type
    // but
    // VAL
    // records
    // coordinates
     val X = 0
 val Y = 1
 val WIDTH = 2
 val HEIGHT = 3

/**
 * returns the encompassing range for this chart, or null if the chart data is
 * too complex to represent
 *
 * @param jsonDataRange
 * @return
 */
     fun getEncompassingDataRange(jsonDataRange:JSONObject):IntArray? {
try
{
val catrange = jsonDataRange.get("c").toString()
val sheet = catrange.substring(0, catrange.indexOf('!'))
val retVals = ExcelTools.getRangeRowCol(catrange)
val nSeries = jsonDataRange.getJSONArray("Series").length()
for (i in 0 until nSeries)
{
val series = jsonDataRange
.getJSONArray("Series").get(i) as JSONObject
val serrange = series.get("v").toString()
if (!serrange.startsWith(sheet))
continue
var locs = ExcelTools.getRangeRowCol(serrange)
try
{
if (locs[0] < retVals[0])
retVals[0] = locs[0]
if (locs[1] < retVals[1])
retVals[1] = locs[1]
if (locs[2] > retVals[2])
retVals[2] = locs[2]
if (locs[3] > retVals[3])
retVals[3] = locs[3]

val legendrange = series.get("l").toString()
locs = ExcelTools.getRowColFromString(legendrange)
if (locs[0] < retVals[0])
retVals[0] = locs[0]
if (locs[1] < retVals[1])
retVals[1] = locs[1]
if (locs[0] > retVals[2])
retVals[2] = locs[0]
if (locs[1] > retVals[3])
retVals[3] = locs[1]

if (series.has("b"))
{
val bubblerange = series.get("b").toString()
locs = ExcelTools.getRangeRowCol(serrange)
if (locs[0] < retVals[0])
retVals[0] = locs[0]
if (locs[1] < retVals[1])
retVals[1] = locs[1]
if (locs[2] > retVals[2])
retVals[2] = locs[2]
if (locs[3] > retVals[3])
retVals[3] = locs[3]
}
}
catch (e:Exception) {
 // just continue
                }

}
return retVals
}
catch (e:Exception) {}

return null
 /*
         * while (ptgs.hasNext()) { PtgRef pr= (PtgRef) ptgs.next();
         * // PtgRef pr=
         * (PtgRef)refs[i]; int[] locs = pr.getIntLocation(); for
         * (int x=0;x<2;x++) {
         * if((locs[x]<retValues[x]))retValues[x]=locs[x]; } for
         * (int x=2;x<4;x++) {
         * if((locs[x]>retValues[x]))retValues[x]=locs[x]; } i++; }
         */
    }

/**
 * Static method to create a new chart on WorkSheet sheet of type chartType with
 * chart Options options <br></br>
 * After creating, you can set the chart title via ChartHandle.setTitle <br></br>
 * and Position via ChartHandle.setRelativeBounds (row/col-based) or
 * ChartHandle.setCoords (pixel-based) <br></br>
 * as well as several other customizations possible
 *
 * @param book      WorkBookHandle
 * @param sheet     WorkSheetHandle
 * @param chartType one of: <br></br>
 * BARCHART <br></br>
 * COLCHART <br></br>
 * LINECHART <br></br>
 * PIECHART <br></br>
 * AREACHART <br></br>
 * SCATTERCHART <br></br>
 * RADARCHART <br></br>
 * SURFACECHART <br></br>
 * DOUGHNUTCHART <br></br>
 * BUBBLECHART <br></br>
 * RADARAREACHART <br></br>
 * PYRAMIDCHART <br></br>
 * CYLINDERCHART <br></br>
 * CONECHART <br></br>
 * PYRAMIDBARCHART <br></br>
 * CYLINDERBARCHART <br></br>
 * CONEBARCHART
 * @param options   EnumSet<ChartOptions>
 * @return
 * @see ChartHandle.ChartOptions
 *
 * @see setChartType
</ChartOptions> */
     fun createNewChart(sheet:WorkSheetHandle, chartType:Int, options:EnumSet<ChartOptions>):ChartHandle {
 // Create Initial Basic Chart
        val cht = sheet.workBook!!.createChart("", sheet)
 // Change Chart Type with Desired Options:
        cht!!.setChartType(chartType, 0, options)
return cht
}
}
}/**
 * Changes or adds a Series to the chart via Series Index. Each bar, line or
 * wedge in a chart represents a Series. <br></br>
 * If the Series index is greater than the number of series already present in
 * the chart, the series will be added to the end. <br></br>
 * Otherwise the Series at the index position will be altered. <br></br>
 * This method allows altering of every aspect of the Series: Data (Series)
 * Range, Legend Text, Legend Cell Address, Category Range and/or Bubble Range.
 *
 * @param int    index - the series index. If greater than the number of series
 * already present in the chart, the series will be added to the end
 * @param String legendCell - String representation of Legend Cell Address
 * @param String legendText - String Legend text
 * @param String categoryRange - String representation of Category Range (should be
 * same for all series)
 * @param String seriesRange - String representation of the Series Data Range for
 * this series
 * @param String bubbleRange - String representation of Bubble Range (representing
 * bubble sizes), if bubble chart. null if not
 * @return a ChartSeriesHandle representing the new or altered Series
 * @throws CellNotFoundException
 */// for
 // default
 // chart
/**
 * Adds a new Series to the chart. Each bar, line or wedge in a chart represents
 * a Series.
 *
 * @param String legendAddress - The cell address defining the legend for the
 * series
 * @param Srring legendText - Text of the legend
 * @param String categoryRange - Cell Range defining the category (normally will be
 * the same range for all series)
 * @param String seriesRange - Cell range defining the data points of the series
 * @param String bubbleRange - Cell range defining the bubble sizes for this series
 * (bubble charts only)
 * @return ChartSeriesHandle representing the new series
 * @throws CellNotFoundException
 */// for
 // default
 // chart
/**
 * remove the Series (bar, line or wedge) at the desired index
 *
 * @param int index - series index (valid values: 0 to
 * getAllChartSeriesHandles().length-1)
 * @see getAllChartSeriesHandles
 */// -1 flag for all series
/**
 * Appends a series one row below the last series in the chart.
 *
 *
 * This can be utilized when programmatically adding rows of data that should be
 * reflected in the chart. <br></br>
 * Legend cell will be incremented by one row if a reference. Category range
 * will stay the same.
 *
 * <br></br>
 * In order for this method to work properly the chart must have row-based
 * series. If your chart utilizes column-based series, then you need to append a
 * category.
 *
 * @return ChartSeriesHandle representing newly added series
 * @see ChartHandle.appendRowCategoryToChart
 */// do for default chart (0)
/**
 * Append a row of categories to the bottom of the chart. <br></br>
 * Expands all Series to include the new bottom row. <br></br>
 * To be utilized when expanding a chart to encompass more data that has a
 * col-based series.
 *
 * @see ChartHandle.appendRowSeriesToChart
 */// default chart
/**
 * makes the default Chart hava a 3D effect <br></br>
 * sets the group of options necessary to create a 3D chart <br></br>
 * For Chart Types: <br></br>
 * BARCHART, COLCHART, LINECHART, PIECHART, AREACHART, BUBBLECHART,
 * PYRAMIDCHART, CYLINDERCHART, CONECHART
 *
 */// bar, col, line, pie, area, bubble, pyramid, cone,
 // cylinder
