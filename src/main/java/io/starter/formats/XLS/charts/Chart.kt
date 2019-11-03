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

import io.starter.OpenXLS.ChartHandle
import io.starter.OpenXLS.ChartHandle.ChartOptions
import io.starter.OpenXLS.ExcelTools
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.XLS.*
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import org.json.JSONArray
import org.json.JSONException

import java.io.*
import java.util.*


/**
 * **Chart: Chart Location and Dimensions 0x1002h**<br></br>
 *
 *
 * The Chart record determines the chart dimensions and
 * marks the beginning of the Chart records.
 *
 *
 * <pre>
 * * Note that all these values are split up 2 bytes integer and 2 bytes fractional
 *
 * offset  name        size    contents
 * ---
 * 4       x           4       x pos of upper left corner
 * 8       y           4       y pos of upper left corner
 * 12      dx          4       x-size
 * 16      dy          4       y-size
 *
</pre> *
 *
 *
 *
 *
 *
 *
 * notes on implementation:
 *
 * <br></br>A chart may have up to 9 chart types (specific instances of a chart e.g. bar, line, etc.) known as chart groups.  These are identified by the chartgroup array list.
 * <br></br>The parent of these chart groups are the axis group.
 * <br></br>A chart's series and trendlines contains indexes into the chart group.
 *
 * @see Ai
 */
open class Chart : GenericChartObject(), ChartObject {

    // Objects which define a chart -- REMEMBER TO UPDATE OOXMLChart copy constructor if modify below member vars
    protected var chartgroup: ArrayList<ChartType> = ArrayList()    // can have up to 4 chart objects (Bar, Area, Line, etc.) per a given chart
    protected var nCharts = 0    // total number of charts (0= the default chart + any overlay charts (up to 4 charts (9 in OOXML) see chartgroup list)
    /**
     * Return the Chart Axes object
     *
     * @return ChartAxis
     */
    var axes: ChartAxes? = null
        protected set
    /**
     * return the Chart Series object
     *
     * @return
     */
    var chartSeries = ChartSeries()
        protected set
    /**
     * return the title TextDisp element
     *
     * @return
     */
    var titleTd: TextDisp? = null
        protected set
    protected var dimensions: Dimensions        // datarange dimensions ...
    // The following two objects are external objects in the XLS stream that are associated with the chart.
    var obj: Obj? = null
        protected set
    var msodrawobj: MSODrawing? = null
        protected set    // stores coordinates and shape info ...
    internal var chartMetrics: HashMap<String, Double> = HashMap()        // hold chartMetrics: x, y, w, h, canvasw, canvash
    @Transient
    var wbh: WorkBookHandle? = null

    // TODO: MANAGE FONTS ---?

    // internal chart records
    protected var chartRecs = ArrayList()
    protected var preRecs: AbstractList<*>? = null
    protected var postRecs: AbstractList<*>? = ArrayList()
    protected var dirtyflag = false    // if anything has changed in the chart (except series, which is handled via another var)
    protected var metricsDirty = true    // initially true so creates min, max and other metrics, true if should be recalculated
    // below vars used to save state in addInitialChartRecord recursion
    protected var currentAi: Ai? = null                // used in init only
    protected var hierarchyDepth = 0    //	""
    protected var initobs = ArrayList()

    /**
     * Returns a map of series ==> series ptgs
     * Ranges can either be in the format "C5" or "Sheet1!C4:D9"
     */
    val seriesPtgs: HashMap<*, *>
        get() = this.chartSeries.seriesPtgs

    /**
     * Returns a serialized copy of this Chart
     */
    val chartBytes: ByteArray?
        get() {
            for (i in preRecs!!.indices) {
                (preRecs!![i] as BiffRec).data
            }

            var obs: ObjectOutputStream? = null
            var b: ByteArray? = null
            try {
                val baos = ByteArrayOutputStream()
                obs = ObjectOutputStream(baos)
                obs.writeObject(this)
                b = baos.toByteArray()
            } catch (e: IOException) {
                Logger.logErr(e)
            }

            return b
        }

    val serialBytes: ByteArray?
        get() {
            for (i in preRecs!!.indices) {
                (preRecs!![i] as BiffRec).data
            }
            var obs: ObjectOutputStream? = null
            var b: ByteArray? = null
            try {
                val baos = ByteArrayOutputStream()
                val bufo = BufferedOutputStream(baos)
                obs = ObjectOutputStream(bufo)
                obs.writeObject(this)
                bufo.flush()
                b = baos.toByteArray()
            } catch (e: IOException) {
                Logger.logErr(e)
            }

            return b
        }

    /**
     * Gets the records for the chart, including the global references, ie the MSODrawing and Obj records
     *
     * @return
     */
    val xlSrecs: List<*>
        get() {
            val l = this.assembleChartRecords()
            if (obj != null) l.add(0, obj)
            if (msodrawobj != null) l.add(0, msodrawobj)
            return l
        }


    /**
     * Dimensions rec holds the datarange for the chart; must update ...
     *
     * @return
     */
    val minRow: Int
        get() = dimensions.rowFirst

    val maxRow: Int
        get() = dimensions.rowLast

    val minCol: Int
        get() = dimensions.colFirst

    val maxCol: Int
        get() = dimensions.colLast

    /**
     * returns the unique Obj Id # of this chart as seen in Vb macros et. al
     *
     * @return in id #
     */
    /**
     * sets the unique id # for this chart
     * used upon addChart
     *
     * @param id
     */
    var id: Int
        get() = if (this.obj != null) this.obj!!.objId else -1
        set(id) {
            if (this.obj != null)
                this.obj!!.objId = id
        }

    /**
     * get all the series objects for ALL charts
     * (i.e. even in overlay charts)
     *
     * @return
     */
    val allSeries: Vector<*>
        get() = chartSeries.getAllSeries(-1)

    /**
     * returns the plot area background color string (hex)
     *
     * @return color string
     */
    val plotAreaBgColor: String
        get() {
            var bg: String? = null
            try {
                val f = Chart.findRec(this.chartArr, Frame::class.java) as Frame?
                bg = f!!.bgColor
            } catch (e: Exception) {
            }

            val ct = this.chartType
            if (bg == null || ct != ChartConstants.PIECHART && ct != ChartConstants.RADARCHART) {
                bg = axes!!.plotAreaBgColor
            }
            if (bg == null)
                bg = "#FFFFFF"
            return bg
        }


    /**
     * returns the plot area line color string (hex)
     *
     * @return color string
     */
    val plotAreaLnColor: String
        get() {
            var clr: String? = "#000000"
            try {
                val f = Chart.findRec(this.chartArr, Frame::class.java) as Frame?
                clr = f!!.lineColor
            } catch (e: Exception) {
            }

            return clr
        }

    /**
     * Get the title of the Chart
     *
     * @return title of the Chart
     */
    /**
     * Set the title of the Chart
     *
     * @param str title of the Chart
     */
    // getTitleTD();
    // No need to create a TD record if setting title to ""
    // if no title and setting title to "", just leave it
    // 20070709 KSC: Adding a new chart, add Title recs ...
    // add TextDisp title to end of chart recs ...
    // just set title text
    var title: String?
        get() = if (titleTd != null) titleTd!!.toString() else ""
        set(str) {
            if ((str == null || str == "") && titleTd == null)
                return
            if (titleTd == null) {
                try {
                    val td = TextDisp.getPrototype(ObjectLink.TYPE_TITLE, str, this.workBook) as TextDisp
                    this.addChartRecord(td)
                    titleTd = td

                } catch (e: Exception) {
                    Logger.logWarn("Unable to set title of chart to: $str This chart object does not contain a title record")
                }

            } else {
                titleTd!!.setText(str)
            }
            setMetricsDirty()
        }


    /**
     * return the default font for the chart
     *
     * @return
     */
    // correct code???? 2 or 3 ??
    val defaultFont: Font?
        get() {
            for (i in chartArr.indices) {
                if ((chartArr[i] as BiffRec).opcode == XLSConstants.DEFAULTTEXT) {
                    if ((chartArr[i] as DefaultText).type == 2) {
                        val td = chartArr[i + 1] as TextDisp
                        val idx = td.fontId
                        return if (idx > -1) wkbook!!.getFont(idx) else null
                    }
                }
            }
            return null
        }

    /**
     * Get all of the Fontx records in this chart.  Utilized for
     * updating font references when moving between workbooks
     */
    val fontxRecs: ArrayList<*>
        get() {
            val ret = ArrayList()
            for (i in chartRecs.indices) {
                val b = chartRecs.get(i) as BiffRec
                if (b.opcode == XLSConstants.FONTX) {
                    ret.add(b)
                }
            }
            return ret
        }


    /**
     * return the font associated with the chart title
     *
     * @return Font
     */
    val titleFont: Font?
        get() = if (titleTd == null) null else titleTd!!.getFont(wkbook!!)

    /**
     * returns the chart type for the default chart
     */
    val chartType: Int
        get() = chartgroup[0].chartType

    /**
     * returns an array of all chart types found in this chart
     * 0= default, 1-9 (max) are overlay charts
     *
     * @return
     */
    // max
    val allChartTypes: IntArray
        get() {
            val charttypes = IntArray(nCharts)
            for (i in 0 until nCharts) {
                charttypes[i] = chartgroup[i].chartType
            }
            return charttypes
        }

    /**
     * returns the number of charts in this chart -
     * 1= default, can have up to 9 overlay charts as well
     *
     * @return
     */
    val numberOfCharts: Int
        get() = nCharts - 1

    /**
     * return the default Chart Object associated with this chart
     *
     * @return
     */
    val chartObject: ChartType
        get() = chartgroup[0]

    /**
     * @return truth of "Chart is Three-D"
     */
    val isThreeD: Boolean
        get() = isThreeD(0)

    /**
     * @return truth of "Chart is Stacked"
     */
    override var isStacked: Boolean
        get() = isStacked(0)
        set(value: Boolean) {
            super.isStacked = value
        }

    /**
     * return truth of "Chart is 100% Stacked"
     *
     * @return
     */
    val is100PercentStacked: Boolean
        get() = is100PercentStacked(0)

    /**
     * @return truth of "Chart is Clustered"  (Bar/Col only)
     */
    val isClustered: Boolean
        get() = isClustered(0)

    /**
     * return ThreeD settings for this chart in XML form
     *
     * @return String XML
     */
    val threeDXML: String
        get() = getThreeDXML(0)

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
    val axisAutomaticScale: Boolean
        get() = this.axes!!.axisAutomaticScale

    /**
     * Return all the Chart-specific font recs in XML form
     * These include the default fonts + title font + axis fonts ...
     *
     * @return
     */
    //((TextDisp) b).getFontId();
    val chartFontRecsXML: String
        get() {
            val fonts = HashMap()
            var maxFont = 5
            for (i in chartRecs.indices) {
                val b = chartRecs.get(i) as BiffRec
                if (b.opcode == XLSConstants.FONTX) {
                    val fontId = (b as Fontx).ifnt
                    maxFont = Math.max(fontId, maxFont)
                    val f = this.workBook!!.getFont(fontId)
                    fonts.put(Integer.valueOf(fontId), f.xml)
                }
            }
            val sb = StringBuffer()
            for (i in 5..maxFont) {
                if (fonts.get(Integer.valueOf(i)) != null) {
                    sb.append("\n\t\t<ChartFontRec id=\"$i\" ")
                    sb.append(fonts.get(Integer.valueOf(i)))
                    sb.append("/>")
                }
            }
            return sb.toString()
        }

    /**
     * return the fontid for non-axis chart fonts (title, default ...)
     *
     * @return
     */
    // should!!
    // should!!
    val chartFontsXML: String
        get() {
            val sb = StringBuffer()
            var i = 0
            while (i < chartArr.size) {
                var b: BiffRec = chartArr[i]
                if (b.opcode == XLSConstants.DEFAULTTEXT) {
                    sb.append(" Default" + (b as DefaultText).type + "=\"")
                    i++
                    b = chartArr[i]
                    if (b.opcode == XLSConstants.TEXTDISP) {
                        val td = b as TextDisp
                        sb.append(td.fontId)
                    }
                    sb.append("\"")
                } else if (b.opcode == XLSConstants.TEXTDISP) {
                    val td = b as TextDisp
                    if (td.isChartTitle) {
                        sb.append(" Title=\"")
                        sb.append(td.fontId)
                        sb.append("\"")
                    }
                }
                i++
            }
            return sb.toString()
        }


    /**
     * return int[4] representing the coordinates (L, T, W, H) of this chart in pixels
     *
     * @return int[4] bounds
     */
    /**
     * set the coordinates of ths chart in pixels
     *
     * @param coords int[4]: L, T, W, H
     */
    // use adjusted values for w and h
    // size is an element in minmax ...
    // should NEVER happen
    var coords: ShortArray
        get() {
            var coords = shortArrayOf(20, 20, 100, 100)
            if (this.msodrawobj != null)
                coords = this.msodrawobj!!.coords
            else
                Logger.logWarn("Chart missing Msodrawing record. Chart has no coordinates.")
            if (!metricsDirty) {
                coords[2] = chartMetrics["canvasw"].toShort()
                coords[3] = chartMetrics["canvash"].toShort()
            }
            return coords
        }
        set(coords) = if (this.msodrawobj != null) {
            msodrawobj!!.coords = coords
            setMetricsDirty()
        } else
            Logger.logWarn("Chart missing coordinates.")

    /**
     * return bounds relative to row/cols including their respective offsets
     *
     * @return
     */
    /**
     * set the bounds relative to row/col and their offsets into them
     *
     * @param bounds short[]
     */
    var bounds: ShortArray?
        get() = if (this.msodrawobj != null) msodrawobj!!.bounds else null
        set(bounds) {
            if (this.msodrawobj != null)
                msodrawobj!!.bounds = bounds
            setMetricsDirty()
        }

    /**
     * returns the offset within the column in pixels
     *
     * @return
     */
    val colOffset: Short
        get() = if (this.msodrawobj != null) msodrawobj!!.colOffset else 0

    /**
     * return the top row of the chart
     */
    val row0: Int
        get() = if (this.msodrawobj != null) msodrawobj!!.row0 else -1

    /**
     * return the left col of the chart
     */
    val col0: Int
        get() = if (this.msodrawobj != null) msodrawobj!!.col else -1

    /**
     * return the height of this chart
     */
    /**
     * set the height of this chart
     *
     * @param h
     */
    var height: Int
        get() = if (this.msodrawobj != null) msodrawobj!!.height.toInt() else -1
        set(h) {
            if (this.msodrawobj != null)
                msodrawobj!!.setHeight(h)
            setMetricsDirty()
        }

    /**
     * retrieve the current Series JSON for comparisons
     *
     * @return JSONArray
     */
    /**
     * set the current Series JSON
     *
     * @param s JSONArray
     */
    var seriesJSON: JSONArray?
        get() = chartSeries.seriesJSON
        @Throws(JSONException::class)
        set(s) {
            chartSeries.seriesJSON = s
            dirtyflag = true
        }

    /**
     * return the Data Legend for this chart, if any
     *
     * @return
     */
    val legend: Legend?
        get() = chartgroup[0].dataLegend

    /**
     * return the default data label setting for the chart, if any
     * <br></br>NOTE: each series can override the default data label fo the chart
     * <br></br>can be one or more of:
     * <br></br>VALUELABEL= 0x1;
     * <br></br>VALUEPERCENT= 0x2;
     * <br></br>CATEGORYPERCENT= 0x4;
     * <br></br>CATEGORYLABEL= 0x10;
     * <br></br>BUBBLELABEL= 0x20;
     * <br></br>SERIESLABEL= 0x40;
     *
     * @return int default data label for chart
     */
    val dataLabel: Int
        get() = chartgroup[0].dataLabel


    /*
    public Chart copy()  throws InstantiationException, IllegalAccessException  {
        Chart copy = this.getClass().newInstance();
        return copy;
     }*/

    override fun init() {
        super.init()
        data
        chartSeries.setParentChart(this)
/**
 * NOTE: the x, y, dx (w) and dy (h) are in most cases
 * not used; the MSODRAWING coords governs the
 * chart canvas; see getMetrics for plot and other adjustments
 *
 * New Info:  have to see if this works
 * Get chart area width in pixels
 * chart area width in pixels = (dx field of Chart record - 8) * DPI of the display device / 72
 * If the frt field of the Frame record following the Chart record is 0x0004 and the chart is not embedded, add the shadow size:
 * chart area width in pixels -= 2 * line width of the display device in pixels
 *
 * Get chart area height in pixels
 * chart area height in pixels = (dy field of Chart record - 8) * DPI of the display device / 72
 * If the frt field of the Frame record following the Chart record is 0x0004 and the chart is not embedded, add the shadow size:
 * chart area height in pixels -= 2 * line height of the display device in pixels
 *
 * NOTE:
 * Since the 1980s, the Microsoft Windows operating system has set the default display "DPI" to 96 PPI,
 * while Apple/Macintosh computers have used a default of 72 PPI.[2]
 *
 * byte[] rkdata = this.getData();
 * try {
 * short s = ByteTools.readShort(rkdata[0],rkdata[1]);
 * short ss = ByteTools.readShort(rkdata[2],rkdata[3]);
 *
 * if(ss<0)ss=0;
 * String parser = s + "." + ss;
 * x = (new Float(parser)).floatValue();
 *
 * s = ByteTools.readShort(rkdata[4],rkdata[5]);
 * ss = ByteTools.readShort(rkdata[6],rkdata[7]);
 * if(ss<0)ss=0;
 *
 * parser = s + "." + ss;
 * y = (new Float(parser)).floatValue();
 *
 *
 * /*Value of the real number = Integral + (Fractional / 65536.0)
 * Integral (2 bytes): A signed integer that specifies the integral part of the real number.
 * Fractional (2 bytes): An unsigned integer that specifies the fractional part of the real number.
 *
 * int integral= 0;
 * integral |= rkdata[8] & 0xFF;
 * integral <<= 8;
 * integral |= rkdata[9] & 0xFF;
 * int fractional= ByteTools.readUnsignedShort(rkdata[10],rkdata[11]);
 * dx= (float)(integral + (fractional/65536.0));
 * integral= ((rkdata[12] & 0xFF) << 8) | (rkdata[13] & 0xFF);
 * if (integral < 0) integral += 256;
 * fractional= ByteTools.readUnsignedShort(rkdata[14],rkdata[15]);
 * dy= (float) integral;
 * } catch (NumberFormatException e) {	// 20080414 KSC: parsing appears to be off - TODO Check out Chart record specs
 * Logger.logErr("Chart.init: parsing dimensions failed: " + e);
 * }
 *
 * if ((DEBUGLEVEL > 3))
 * Logger.logInfo("Chart Found @ x:" + x + " y:" + y + " dx:" + dx + " dy:" + dy); */
}


/*
 * Note, other important chart records we are not as of yet trapping:
 * CttMlFrt 0x89E -- additional properties for chart elements
 * ShapePropsStream	0x8A4
 * TextPropsStream	0x8A5
 * CrtLayout12  	0x89D	-- layout info for attached label, legend
 * 0x89E=The CrtMlFrt record specifies additional properties for chart elements, as specified by the Chart Sheet SubstreamABNF. These properties complement the record to which they correspond, and are stored as a structure chain defined in XmlTkChain. An application can ignore this record without loss of functionality, except for the additional properties. If this record is longer than 8224 bytes, it MUST be split into several records. The first section of the data appears in this record and subsequent sections appear in one or more CrtMlFrtContinue records that follow this record.
 *
 * Legend --> Pos, TextDisp, Frame, CrtLayout12, TextPropsStream, CrtLayout12
*/

/**
 * Adds chart records to the chartRecs array.  This array should be completed with all chart records
 * before the initChartRecords method is called, creating the hierarchical chart structure.
 *
 * @param br - a chart record
 * @return false if we reach the final "end" record, signifying the end of the chart object structure.
 *
 *
 * Be aware that there can be additional chart records after this structure,  Chart only deals with it's own heirarchy.
*/
fun addInitialChartRecord(br:BiffRec):Boolean {
if (br.opcode.toInt() == 0x1033)
{
++hierarchyDepth
}
else if (br.opcode.toInt() == 0x1034)
{
--hierarchyDepth
if (hierarchyDepth == 0)
{
chartRecs.add(br)
return false
}
}
if (hierarchyDepth != 0)
{
if (br.opcode == XLSConstants.AI)
{
currentAi = br as Ai
currentAi!!.parentChart = this  // necessary for base ops
currentAi!!.setSheet(this.sheet) // needed for cell change updates
}
else if (currentAi != null && br.opcode == XLSConstants.SERIESTEXT)
{
currentAi!!.setSeriesText(br as SeriesText)
currentAi = null    // done
}
else if (br.opcode == XLSConstants.SERIES)
{
chartSeries.add(arrayOf<Any>(br, Integer.valueOf(nCharts)))    // store series records - initially map all to default chart (normal case)
}
else if (br.opcode == XLSConstants.AXISPARENT)
{
axes = ChartAxes(br as AxisParent)
}
else if (br.opcode == XLSConstants.AXIS)
{
axes!!.add(br as Axis)
}
else if (br.opcode == XLSConstants.CHARTFORMAT)
{    // usually only 1 (default chart id= 0) but can have up to 9 overlay charts
nCharts++
}
else if (br.opcode == XLSConstants.SERIESLIST)
{ // maps chart # to series #
chartSeries.addSeriesMapping(nCharts - 1, (br as SeriesList).seriesMappings!!)// 1-based series #'s to "assign" to the overlay chart
}
try
{
if ((br as GenericChartObject).chartType != -1)
{ // store the chart object which defines this chart (NOTE: may be up to 4 charts in one = overlay charts)
var cf:ChartFormat? = null
for (i in chartRecs.indices.reversed())
{
if ((chartRecs.get(i) as BiffRec).opcode == XLSConstants.CHARTFORMAT)
{
cf = chartRecs.get(i) as ChartFormat
break
}
}
// sadly, can't do it here - have to wait until process hierarchy chartobj.add(ChartType.getChartObject((GenericChartObject)br, cf, this.getWorkBook()));
initobs.add(arrayOf<BiffRec>(br, cf))
} // Set Parent Chart as it's necessary for all basic chart ops
(br as GenericChartObject).parentChart = this  // Note: Scl is not a GenericChartObject + can have unknown (XLSRecords) - will cause ClassCastException
}
catch (c:ClassCastException) {}

chartRecs.add(br)
}
else
{
if (br.opcode == XLSConstants.DIMENSIONS) dimensions = br as Dimensions
br.data
postRecs!!.add(br)
}

return true
}


fun setDirtyFlag(b:Boolean) {
dirtyflag = b
}

/**
 * Take the initial array of records for the chart, and create a
 * hierarchial array of objects for the chart.  Also, populate initial values for easy access.
*/
fun initChartRecords() {
try
{    // Added to find/set obj + msodrawing records - replaces setObj
// see TestReadWrite.TestIOOBError
// For Normal charts, Obj rec is the last record before the Chart record
if (!this.sheet!!.isChartOnlySheet)
{
var pos = this.sheet!!.sheetRecs.size - 1
var rec = this.sheet!!.sheetRecs.get(pos) as BiffRec
obj = rec as Obj
obj!!.chart = this
// Usually, MsoDrawing is just preceding Obj record, except in those rare carses where
// there are Continues and Txo's ...
while (--pos > 0)
{
rec = this.sheet!!.sheetRecs.get(pos) as BiffRec
if (rec.opcode == XLSConstants.MSODRAWING)
{
this.msodrawobj = rec as MSODrawing
break
}
}
} // chart-only worksheets have no obj/mso apparently
}
catch (e:Exception) {
Logger.logErr("initChartRecords: Error in Chart Records:  " + e.toString())
}

// Turn it into a static array initially to speed up random access
var bArr = arrayOfNulls<BiffRec>(chartRecs.size)
bArr = chartRecs.toTypedArray() as Array<BiffRec>
this.initChartObject(this, bArr)
for (i in initobs.indices)
{
val ios = initobs.get(i) as Array<BiffRec>
chartgroup.add(ChartType.createChartTypeObject(ios[0] as GenericChartObject, ios[1] as ChartFormat, this.workBook))
val l = Chart.findRec((ios[1] as ChartFormat).chartArr, Legend::class.java) as Legend?
if (l != null)
chartgroup.get(chartgroup.size - 1).addLegend(l)
}
initobs = ArrayList()    // clear out
}

/**
 * Handle the iteration and creation of the chart objects
*/
private fun initChartObject(cobj:ChartObject, cRecs:Array<BiffRec>):ChartObject {

var i = 0
while (i < cRecs.size)
{
val b = cRecs[i]
b.data
if (!(b.opcode == XLSConstants.BEGIN) && !(b.opcode == XLSConstants.END))
{
if ((cRecs.size > i + 1) && ((cRecs[i + 1]).opcode == XLSConstants.BEGIN))
{ // this is an object with sub-data
try
{
val co = b as ChartObject
val endloc = this.getMatchingEndRecordLocation(i, cRecs)
val arrlen = endloc - i
val objArr = arrayOfNulls<BiffRec>(arrlen)
System.arraycopy(cRecs, i + 1, objArr, 0, arrlen)
cobj.addChartRecord(this.initChartObject(co, objArr) as XLSRecord)
// necessary initialization of key elements
if (co is TextDisp)
{
val type = (co as TextDisp).type
// -1 = default text ...?
if (type == ObjectLink.TYPE_TITLE)
titleTd = co as TextDisp
else if (type == ObjectLink.TYPE_XAXIS)
axes!!.setTd(ChartConstants.XAXIS, co as TextDisp)
else if (type == ObjectLink.TYPE_YAXIS)
axes!!.setTd(ChartConstants.YAXIS, co as TextDisp)
else if (type == ObjectLink.TYPE_ZAXIS)
axes!!.setTd(ChartConstants.ZAXIS, co as TextDisp)
else if (type == ObjectLink.TYPE_DATAPOINTS)

else if (type == ObjectLink.TYPE_DISPLAYUNITS)
;//
// KSC: TESTING!! Take out when done
//                            	Logger.logInfo("Display Units");
// series or data points
// do what??
}
try
{
co.parentChart = this
}
catch (e:Exception) {}

i += arrlen
}
catch (e:ClassCastException) {
// it's not a defined chart object.  Add it in!!!  If we are missing a chart object containing other records we
// will not be able to write these out correctly.
Logger.logWarn("Error in parsing chart. Please add the correct object (opcode: " + b.opcode + ") to be a Chart Object")
}

}
else
{
cobj.addChartRecord(b as XLSRecord)
}
}
i++
}
return cobj
}

/**
 * Just a little helper method!  Determines where the matching end record is for a begin record
 *
 * @param startLoc
 * @param cRecs
 * @return
*/
private fun getMatchingEndRecordLocation(startLoc:Int, cRecs:Array<BiffRec>):Int {
var offset = 0
for (i in startLoc + 2 until cRecs.size)
{
val b = cRecs[i]
if (b.opcode == XLSConstants.BEGIN) offset++
if (b.opcode == XLSConstants.END)
{
offset--
if (offset < 0) return i
}
}
return -1
}

/**
 * put together the chart records for output to the streamer
 *
 * @return arranged List of Chart Records.
*/
fun assembleChartRecords():MutableList<*> {
val outputVec = Vector()
if (preRecs != null) outputVec.addAll(preRecs!!)
if (dirtyflag)
{
for (i in 0 until nCharts)
{
chartSeries.updateSeriesMappings(chartgroup.get(i).seriesList, i)    // if has multiple or overlay charts, update series mappings
}
outputVec.addAll(this.recordArray)
}
else
{
outputVec.add(this)
outputVec.addAll(chartRecs)
}
if (postRecs != null) outputVec.addAll(postRecs!!)
if (!true)
{
for (i in outputVec.indices)
{
val rec = outputVec.get(i) as XLSRecord
Logger.logInfo("rec:" + rec.toString() + "")
Logger.logInfo(ByteTools.getByteDump(rec.getData(), 0))
Logger.logInfo("-------------------------------------------------")
}
}
return outputVec
}

/**
 * This handles setting the records external to, but supporting the chart object hierarchy.
 *
 *
 * Think of this kind of like the recvec & cell records in the XLS parsing we do.
 *
 *
 * We have a pre and post array of biffrecs that are (for now)  unmodifieable.
 * These get appended into the whole array on ouput in append chart records.
 *
 * @param recs
*/
fun setPreRecords(recs:AbstractList<*>) {
preRecs = recs
}

fun setDimensionsRecord(r0:Int, r1:Int, c0:Int, c1:Int) {
dimensions.setRowFirst(r0 - 1)
dimensions.setRowLast(r1 - 1)
dimensions.setColFirst(c0)
dimensions.setColLast(c1 - 1)
/* also must remove label and numberrec cached records if altered series
otherwise causes errors when removing or altering series
*/
for (i in postRecs!!.size - 1 downTo 1)
{
val b = postRecs!!.get(i) as BiffRec
val op = b.opcode.toInt()
if (op == XLSConstants.NUMBER.toInt() || op == XLSConstants.LABEL.toInt())
{
postRecs!!.removeAt(i)
}
}
}


fun setDimensions(d:Dimensions) {
dimensions = d
}

fun setDimensionsRecord() {
val serieslist = this.getAllSeries(-1)
val nSeries = serieslist.size
var nPoints = 0
for (i in serieslist.indices)
{
try
{
val s = serieslist.get(i) as Series
val coords = ExcelTools.getRangeCoords(s.seriesValueAi!!.definition)
if (coords[3] > coords[1])
nPoints = Math.max(nPoints, coords[3] - coords[1] + 1)    // c1-c0
else
nPoints = Math.max(nPoints, coords[2] - coords[0] + 1)    // r1-r0
}
catch (e:Exception) {}

}
this.setDimensionsRecord(0, nPoints, 0, nSeries)
}

/**
 * flags chart metrics should
 * to be recalculated
*/
fun setMetricsDirty() {
dirtyflag = true
metricsDirty = true
}

/**
 * Change series ranges for ALL matching series
 *
 * @param originalrange
 * @param newrange
 * @return
*/
fun changeSeriesRange(originalrange:String, newrange:String):Boolean {
setMetricsDirty()
return chartSeries.changeSeriesRange(originalrange, newrange)
}

/**
 * Return a string representing all series in this chart
 *
 * @param nChart 0=default, 1-9= overlay charts -1 for ALL charts
 * @return
*/
fun getSeries(nChart:Int):Array<String> {
return chartSeries.getSeries(nChart)
}

/**
 * Return an array of strings, one for each category
 *
 * @param nChart 0=default, 1-9= overlay charts -1 for All
 * @return
*/
fun getCategories(nChart:Int):Array<String> {
return chartSeries.getCategories(nChart)
}

/**
 * get all the series objects in the specified chart (-1 for ALL)
 *
 * @param nChart 0=default, 1-9= overlay charts -1 for ALL series
 * @return
*/
fun getAllSeries(nChart:Int):Vector<*> {
return chartSeries.getAllSeries(nChart)
}

/**
 * Add a series object to the array.
 *
 * @param seriesRange   = one row range, expressed as (Sheet1!A1:A12);
 * @param categoryRange = category range
 * @param bubbleRange=  bubble range, if any 20070731 KSC
 * @param seriesText    = label for the series;
 * @param nChart        0=default, 1-9= overlay charts
 * s     * @return
*/
fun addSeries(seriesRange:String, categoryRange:String, bubbleRange:String, legendRange:String, legendText:String, nChart:Int):Series {
val s = chartSeries.addSeries(seriesRange, categoryRange, bubbleRange, legendRange, legendText, chartgroup.get(nChart), nChart)
setMetricsDirty()
return s
}


/**
 * return an array of legend text
 *
 * @param nChart 0=default, 1-9= overlay charts -1 for ALL
*/
fun getLegends(nChart:Int):Array<String> {
return chartSeries.getLegends(nChart)
}

/**
 * specialty method to take absoulte index of series and remove it
 * <br></br>only used in WorkSheetHandle
 *
 * @param index
*/
fun removeSeries(index:Int) {
setMetricsDirty()
chartSeries.removeSeries(index)
}

/**
 * remove desired series from chart
 *
 * @param index
*/
fun removeSeries(seriestodelete:Series) {
// remove from cache
setMetricsDirty()
chartSeries.removeSeries(seriestodelete)
// remove from chartArray
var nSeries = -1
for (i in chartArr.indices)
{
var b:BiffRec = chartArr.get(i)
if (b.opcode == XLSConstants.SERIES)
{
nSeries++
if (seriestodelete == b)
{    // this is the one to delete
chartArr.removeAt(i)
// now adjust series number for all subsequent Series
for (j in i until chartArr.size)
{
b = chartArr.get(j)
if (b.opcode == XLSConstants.SERIES)
{
val s = b as Series
val x = Chart.findRecPosition(s.chartArr, DataFormat::class.java)
if (x > 0)
{
val df = s.chartArr.get(x) as DataFormat
df.setSeriesIndex(nSeries++)
}
}
else
// we've got 'em all
break
}
// now make sure referenced label and number recs are removed
/** how does this work, it does not reference a particular series label, just all of them?
 * for (int j= 0; j < postRecs.size(); j++) {
 * if ((postRecs.get(j) instanceof NumberRec) ||
 * (postRecs.get(j) instanceof Label)) {
 * postRecs.remove(j);
 * j--;
 * }
 *
 * }
*/
break
}
}
}
}

/**
 * get a chart series handle based off the series name
 *
 * @param seriesName
 * @param nChart     0=default, 1-9= overlay charts
 * @return
*/
fun getSeries(seriesName:String, nChart:Int):Series? {
return chartSeries.getSeries(seriesName, nChart)
}

/**
 * @param originalrange
 * @param newrange
 * @return
*/
fun changeCategoryRange(originalrange:String, newrange:String):Boolean {
setMetricsDirty()
return chartSeries.changeCategoryRange(originalrange, newrange)
}


fun changeTextValue(originalval:String, newval:String):Boolean {
setMetricsDirty()
return chartSeries.changeTextValue(originalval, newval)
}

/**
 * sets the plot area background color
 *
 * @param bg color int
*/
fun setPlotAreaBgColor(bg:Int) {
this.axes!!.setPlotAreaBgColor(bg)
setMetricsDirty()
}

/**
 * return the textdisp object that defines the chart title
 *
 * @return public TextDisp getTitleTD(){
 * if(title != null){
 * return title;
 * }
 * for (int i=0;i<chartArr.size></chartArr.size>();i++) {
 * BiffRec b = (BiffRec)chartArr.get(i);
 * if (b.getOpcode()==TEXTDISP) {
 * TextDisp td = (TextDisp)b;
 * if (td.isChartTitle()){
 * title = td;
 * }
 * }
 * }
 * return title;
 * }
*/

public override fun toString():String {
val t = title
if (t != "")
return t
return "Untitled Chart"
}

/**
 * Set the default font for the specific DefaultText rec
 *
 * @param type
 * @param fontId
*/
fun setDefaultFont(type:Int, fontId:Int) {
for (i in chartArr.indices)
{
val b = chartArr.get(i)
if (b.opcode == XLSConstants.TEXTDISP)
{
//xxxxx
}
}
dirtyflag = true
}

/**
 * resets all fonts in the chart to the default font of the workbook Useful
 * for OOXML charts which are based upon default chart
*/
fun resetFonts() {
for (i in chartRecs.indices)
{
val b = chartRecs.get(i) as BiffRec
if (b.opcode == XLSConstants.FONTX)
{
(b as Fontx).setIfnt(0)
}
}
}

/**
 * Utilized by external chart copy, this
 * sets references internally in this chart which allows cross-workbook
 * identification of worksheets
 * to occur.
*/
fun populateForTransfer() {
val i = this.getAllSeries(-1).iterator()
while (i.hasNext())
{
val s = i.next() as Series
try
{
for (j in s.chartArr.indices)
{
if ((s.chartArr.get(j) as BiffRec).opcode == XLSConstants.AI)
{
(s.chartArr.get(j) as Ai).populateForTransfer(this.sheet!!.sheetName)
}
}
}
catch (e:Exception) {
Logger.logErr("Chart.populateForTransfer: " + e.toString())
}

}

}

/**
 * update Sheet References contained in associated Series/Category Ai's
 * using saved Original Sheet References
*/
fun updateSheetRefs(newSheetName:String, origWorkBook:String) {
val i = this.getAllSeries(-1).iterator()
while (i.hasNext())
{
val s = i.next() as Series
try
{
for (j in s.chartArr.indices)
{
if ((s.chartArr.get(j) as BiffRec).opcode == XLSConstants.AI)
{
(s.chartArr.get(j) as Ai).updateSheetRef(newSheetName, origWorkBook)    // 20080630 KSC: add sheet name for lookup
}
}
}
catch (e:Exception) {
Logger.logErr("Chart.updateSheetRefs: " + e.toString())
}

}
}

/**
 * returns the chart type for the desired chart
 *
 * @param nChart 0=default, 1-9= overlay charts
 * @return
*/
fun getChartType(nChart:Int):Int {
return chartgroup.get(nChart).chartType
}

/**
 * return the nth Chart Object associated with this chart (default=0, overlay=1 thru 9)
 *
 * @return
*/
fun getChartObject(nChart:Int):ChartType {
return chartgroup.get(nChart)
}


/**
 * gets the integral order of the specific chart obect in question
 *
 * @param ct Chart Type rwobject
 * @return
*/
fun getChartOrder(ct:ChartType):Int {
for (i in chartgroup.indices)
{
if (chartgroup.get(i) == ct)
return i
}
return 0
}

/**
 * create internal records necessary for an overlay chart
*/
private fun createNewChart(nChart:Int):ChartFormat? {
val ap = findRec(chartArr, AxisParent::class.java) as AxisParent?
return ap!!.createNewChart(nChart)
}


/**
 * adds a new Chart Type to the group of Chart Type Objects, or replaces one at existing index
 *
 * @param ct     Chart Type Object
 * @param nChart index
*/
fun addChartType(ct:ChartType, nChart:Int) {
if (nChart < chartgroup.size)
chartgroup.removeAt(nChart)
else
nCharts++
chartgroup.add(nChart, ct)
dirtyflag = true
}

/**
 * retrieves the parent of the chart type object (==ChartFormat) at index nChart, or creates a new Chart Type Parent
 *
 * @param nChart
 * @return ChartFormat
*/
fun getChartOjectParent(nChart:Int):ChartFormat? {
var cf:ChartFormat? = null
if (nChart >= chartgroup.size)
// create new
cf = this.createNewChart(nChart)
else
cf = chartgroup.get(nChart).cf
return cf
}

/**
 * changes this chart to be a specific chart type with specific display options
 * <br></br>for multiple charts, specify nChart 1-9, for the default chart,
 * nChart= 0
 *
 * @param chartType             chart type one of: BARCHART, LINECHART, AREACHART, COLCHART, PIECHART,
 * DONUGHTCHART, RADARCHART, RADARAREACHART, PYRAMIDCHART, CONECHART, CYLINDERCHART,
 * SURFACTCHART
 * @param nChart                chart # 0 is default
 * @param EnumSet<ChartOptions> 0 or more chart options (Such as Stacked, Exploded ...)
 * @see ChartOptions
</ChartOptions> */
fun setChartType(chartType:Int, nChart:Int, options:EnumSet<ChartOptions>) {
val c:GenericChartObject

val cf = getChartOjectParent(nChart)
val ct = ChartType.create(chartType, this, cf)
ct!!.setOptions(options)

// save and reset legend:
val l = this.legend
ct!!.addLegend(l)

if (ct is BubbleChart && options.contains(ChartOptions.THREED))
(ct as BubbleChart).is3d = true        // when set, every series created will

addChartType(ct, nChart)


/* TODO: axes, other chart options ... Legend???
// The axis group MUST contain two value axes if and only if all chart groups are of type bubble or scatter.
} else if (chartType==BUBBLECHART || chartType==SCATTERCHART) {
// TODO: FINISH!!! IS NOT CORRECT UNLESS PROPER AXES ARE CREATED
//The axis group MUST contain a category or date axis if the axis group contains an area, bar, column, filled radar, line, radar, or surface chart group
} else { // ensure has proper axes
// TODO: FINISH!!! IS NOT CORRECT UNLESS PROPER AXES ARE CREATED
}
*/
/**
 * The axis group MUST contain a series axis if and only if the chart group attached to the axis group is one of the following:
 * An area chart group with the fStacked field of the Area record equal to 0.
 * A column chart group with the fStacked field of the Bar record equal to 0 and the fClustered field of the Chart3d record equal to 0.
 * A line chart group with field fStacked of the Line record equal to 0.
 * A surface chart group
 * The chart group on the axis group MUST contain a Chart3d record if the axis group contains a series axis.
*/
/**
 * for overlay charts i.e. multiple charts, restrictions:
 * Because there are many different ways to represent data visually, each representation has specific requirements about the layout of the data and the way it is plotted.
 * This results in restrictions on the combinations of chart group types that can be plotted on the same axis group, and the combinations of chart group types that can
 * be plotted in the same chart.
 * A chart MUST contain one of the following:
 * A single axis group that contains a single chart group that contains a Chart3d record.
 * One or two axis groups that each contain a single bubble chart group.
 * One or two axis groups that each conform to one of the following restrictions on chart group type combinations:
 * Zero or one of each of the following chart group types: area, column, line, and scatter.
 * Zero or one of each of the following chart group types: bar of pie, doughnut, pie, and pie of pie.
 * A single bar chart group.
 * A single filled radar chart group.
 * A single radar chart group.
 *
*/
}

/**
 * @param nChart 0=default, 1-9= overlay charts
 * @return truth of "Chart is Three-D"
*/
fun isThreeD(nChart:Int):Boolean {
return chartgroup.get(nChart).isThreeD
}

/**
 * return the 3d rec for this chart or null if it doesn't exist
 *
 * @param nChart 0=default, 1-9= overlay charts
 * @return
*/
fun getThreeDRec(nChart:Int):ThreeD? {
return chartgroup.get(nChart).getThreeDRec(false)
}

/**
 * @param nChart 0=default, 1-9= overlay charts
 * @return truth of "Chart is Stacked"
*/
fun isStacked(nChart:Int):Boolean {
return chartgroup.get(nChart).isStacked
}

/**
 * return truth of "Chart is 100% Stacked"
 *
 * @param nChart 0=default, 1-9= overlay charts
 * @return
*/
fun is100PercentStacked(nChart:Int):Boolean {
return chartgroup.get(nChart).is100PercentStacked
}

/**
 * @param nChart 0=default, 1-9= overlay charts
 * @return truth of "Chart is Clustered"  (Bar/Col only)
*/
fun isClustered(nChart:Int):Boolean {
return chartgroup.get(nChart).isClustered
}

/**
 * return chart-type-specific options in XML form
 *
 * @param nChart 0=default, 1-9= overlay charts
 * @return String XML
*/
fun getChartOptionsXML(nChart:Int):String {
return chartgroup.get(nChart).chartOptionsXML
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
public override fun setChartOption(op:String, `val`:String):Boolean {
dirtyflag = true
return this.setChartOption(op, `val`, 0)
}

/**
 * interface for setting chart-type-specific options
 * in a generic fashion
 *
 * @param op     option
 * @param val    value
 * @param nChart 0=default, 1-9= overlay charts
 * @see OpenXLS.handleChartElement
 *
 * @see ChartHandle.getXML
*/
fun setChartOption(op:String, `val`:String, nChart:Int):Boolean {
dirtyflag = true
return chartgroup.get(nChart).setChartOption(op, `val`)
}

/**
 * get the value of *almost* any chart option (axis options are in Axis)
 * for the default chart
 *
 * @param op String option e.g. Shadow or Percentage
 * @return String value of option
*/
public override fun getChartOption(op:String):String? {
return chartgroup.get(0).getChartOption(op)
}

/**
 * return ThreeD settings for this chart in XML form
 *
 * @param nChart 0=default, 1-9= overlay charts
 * @return String XML
*/
fun getThreeDXML(nChart:Int):String {
return chartgroup.get(nChart).threeDXML
}

/**
 * returns the 3d record of the desired chart
 *
 * @param nChart    0=default, 1-9= overlay charts
 * @param chartType one of the chart type constants
 * @return
*/
@JvmOverloads  fun initThreeD(nChart:Int = 0, chartType:Int = this.getChartType(0)):ThreeD {
return chartgroup.get(nChart).initThreeD(chartType)
}

/**
 * returns the minimum and maximum values by examining all series on the chart
 *
 * @param bk
 * @return double[] min, max
*/
fun getMinMax(wbh:WorkBookHandle):DoubleArray {
if (metricsDirty)
getMetrics(wbh)
if (this.wbh == null)
this.wbh = wbh
metricsDirty = false
val minmaxcache = chartSeries.getMetrics(metricsDirty)    // Ignore Overlay charts for now!
return minmaxcache
}

/**
 * Set the specific chart font (title, default ...)
 * NOTE: Axis fonts are handled separately
 *
 * @param type
 * @param val
 * @see Axis.setChartOption, TextDisp.setChartOption
*/
fun setChartFont(type:String, `val`:String) {
var type = type
if (type.equals("Title", ignoreCase = true))
for (i in chartArr.indices)
{
val b = chartArr.get(i)
if (b.opcode == XLSConstants.TEXTDISP)
{    // should be the title
val td = b as TextDisp
if (td.isChartTitle)
{        // should!!
td.fontId = Integer.parseInt(`val`)
break
}
}
}
else if (type.indexOf("Default") > -1)
{
type = type.substring(7)
var i = 0
while (i < chartArr.size)
{
var b:BiffRec = chartArr.get(i)
if (b.opcode == XLSConstants.DEFAULTTEXT)
{
if ((b as DefaultText).type == Integer.parseInt(type))
{
i++
b = chartArr.get(i)
if (b.opcode == XLSConstants.TEXTDISP)
{    // should be!!
val td = b as TextDisp
td.fontId = Integer.parseInt(`val`)
break
}
}
}
i++
}
}
setMetricsDirty()
}


/**
 * set the fontId for the chart title rec
 *
 * @param fontId
*/
fun setTitleFont(fontId:Int) {
for (i in chartArr.indices)
{
val b = chartArr.get(i)
if (b.opcode == XLSConstants.TEXTDISP)
{
val td = b as TextDisp
if (td.isChartTitle)
{
td.fontId = fontId
}
}
}
setMetricsDirty()
}


// 20070802 KSC: debugging utility to write out chart recs
fun writeChartRecs(fName:String) {
try
{
val f = java.io.File(fName)
val writer = BufferedWriter(FileWriter(f, true))
var ctr = 0
if (preRecs != null)
ctr = ByteStreamer.writeRecs(ArrayList(preRecs!!), writer, ctr, 0)
ctr = ByteStreamer.writeRecs(chartArr, writer, ctr, 0)    // will recurse
writer.flush()
writer.close()
}
catch (e:Exception) {}

}

/**
 * set the sheet for this chart plus its subrecords as well
*/
public override fun setSheet(b:Sheet?) {
super.setSheet(b)
if (this.msodrawobj != null)
this.msodrawobj!!.setSheet(b)
for (i in chartArr.indices)
{
chartArr.get(i).setSheet(b)
}
}

/**
 * return the coordinates of the outer plot area (plot + axes + labels) in pixels
 *
 * @return
*/
fun getPlotAreaCoords(w:Float, h:Float):FloatArray? {
return this.axes!!.getPlotAreaCoords(w, h)
}

/**
 * set the top row for this chart
 *
 * @param r
*/
fun setRow(r:Int) {
if (this.msodrawobj != null)
msodrawobj!!.setRow(r)
setMetricsDirty()
}

/**
 * Show or remove Data Table for Chart
 * NOTE:  METHOD IS STILL EXPERIMENTAL
 *
 * @param bShow
*/
fun showDataTable(bShow:Boolean) {
// TODO: FINISH		chartobj[0].showDataTable(bShow);
}

/**
 * show or hide chart legend key
 *
 * @param bShow    boolean show or hide
 * @param vertical boolean show as vertical or horizontal
*/
fun showLegend(bShow:Boolean, vertical:Boolean) {
chartgroup.get(0).showLegend(bShow, vertical)
}

/**
 * remove the legend from the chart
*/
open fun removeLegend() {
showLegend(false, false)
}

/**
 * return data label options for each series as an int array
 * <br></br>each can be one or more of:
 * <br></br>VALUELABEL= 0x1;
 * <br></br>VALUEPERCENT= 0x2;
 * <br></br>CATEGORYPERCENT= 0x4;
 * <br></br>CATEGORYLABEL= 0x10;
 * <br></br>BUBBLELABEL= 0x20;
 * <br></br>SERIESLABEL= 0x40;
 *
 * @return int array
 * @see AttachedLabel
*/
fun getDataLabelsPerSeries(defaultDL:Int):IntArray {
/* NOTES:		 *
 * A data label is a label on a chart that is associated with a data point, or associated with a series  on an area or filled radar chart group.
 * A data label contains information about the associated data point, such as the description of the data point, a legend key, or custom text.
 *
 * Inheritance
 * For any given data point, there is an order of inheritance that determines the contents of a data label associated with the data point:
Data labels can be specified for a chart group, specifying the default setting for the data labels associated with the data points on the chart group .
Data labels can be specified for a series, specifying the default setting for the data labels associated with the data points of the series.
This type of data label overrides the data label properties specified on the chart group for the data labels associated with the data points in a given series.
Data labels can be specified for a data point, specifying the settings for a data label associated with a particular data point.
This type of data label overrides the data label properties specified on the chart group and series for the data labels associated with a given data point.

 * If formatting is not specified for an individual data point, the data point inherits the formatting of the series.
 * If formatting is not specified for the series, the series inherits the formatting of the chart group that contains the series.
 * The yi field of the DataFormat record MUST specify the zero-based index of the Series record associated with this series in the
 * collection of all Series records in the current chart sheet substream that contains the series.
*/
return chartSeries.getDataLabelsPerSeries(defaultDL, this.chartType)
}

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
protected fun getMarkerFormats(nChart:Int):IntArray {
val mf = chartSeries.markerFormats
for (marker in mf)
{
if (marker != 0)
return mf
}
// see if chart format
return chartgroup.get(nChart).markerFormats
}

/**
 * returns true if this chart has data markers (line, scatter and radar charts only)
 *
 * @return
*/
fun hasMarkers(nChart:Int):Boolean {
val markers = this.getMarkerFormats(nChart)
for (marker in markers)
{
if (marker != 0)
return true
}
return false

}

/**
 * return truth if Chart has a data legend key showing
 *
 * @return
*/
fun hasDataLegend():Boolean {
return (chartgroup.get(0).dataLegend != null)
}

/**
 * sets bar colors to vary or not
 *
 * @param vary
 * @param nChart
*/
fun setVaryColor(vary:Boolean, nChart:Int) {
this.getChartObject(nChart).cf!!.varyColor = vary
}

// TODO: LATER
// **** INCLUDE cell selection in series bars ...
// chart.getSeries vs chart.getSeries(int) -- rename!!!
// chart axes need a dirty flag to rebuild metrics
// Axis.getSVG ==> passed categories -- should be passed chartseries instead??
// merge minmaxcache into chartMetrics
// metricsDirty - ensure all ops are covered ...
// clean up Axes.getSVG w.r.t. label font, etc.
fun getMetrics(wbh:WorkBookHandle):HashMap<*, *> {
if (metricsDirty)
{
try
{
this.wbh = wbh
//chartseries.setWorkBook(wbh);
val minmax = chartSeries.getMetrics(metricsDirty)    // Ignore Overlay charts for now!
val coords = this.coords
chartMetrics.put("x", coords[0])
chartMetrics.put("y", coords[1])
chartMetrics.put("w", coords[2])
chartMetrics.put("h", coords[3])
chartMetrics.put("canvasw", coords[2])
chartMetrics.put("canvash", coords[3])
chartMetrics.put("min", minmax[0])
chartMetrics.put("max", minmax[1])
var plotcoords:FloatArray? = null
plotcoords = this.getPlotAreaCoords(chartMetrics.get("w").toFloat(), chartMetrics.get("h").toFloat())
if (plotcoords == null)
{
val crt = Chart.findRec(this.chartArr, CrtLayout12A::class.java) as CrtLayout12A?
if (crt != null)
plotcoords = crt!!.getInnerPlotCoords(chartMetrics.get("w").toFloat(), chartMetrics.get("h").toFloat())
}

chartMetrics.put("x", plotcoords!![0])
chartMetrics.put("y", plotcoords!![1])
chartMetrics.put("w", plotcoords!![2])
chartMetrics.put("h", plotcoords!![3])
// Chart title offset
val titlefont = this.titleFont
if (titlefont != null && this.title != "")
{ // apparently can still have td even when no title is present ...
val tdcoords = this.titleTd!!.coords
val fh = titlefont!!.fontHeightInPoints
if (tdcoords!![1] == 0f)
{
chartMetrics.put("TITLEOFFSET", Math.ceil(fh * 1.5))    // with padding
}
else
{
chartMetrics.put("TITLEOFFSET", fh)    // a little padding
}
}
else if (chartMetrics.get("y") < 5.0)
chartMetrics.put("TITLEOFFSET", 10.0)    // no title offset - a little padding
else
chartMetrics.put("TITLEOFFSET", 0.0)    // no title offset and no need for padding
this.axes!!.getMetrics(this.chartType, chartMetrics, plotcoords, this.chartSeries.getCategories())
var lcoords:IntArray? = null
var adjust = 10.0
if (this.legend != null)
{
this.legend!!.getMetrics(chartMetrics, this.chartType, this.chartSeries)
lcoords = this.legend!!.coords
if (lcoords != null)
// TODO: legend adjustment may have to do with y title and label ofsets ...?
{
adjust = (2 * lcoords!![4]).toDouble()    // spacing before and after legend box  TODO this isn't correct !!
//KSC: TESTING!
//io.starter.toolkit.Logger.log("Original lcoords:  " + Arrays.toString(lcoords));
}
else
{
lcoords = IntArray(6)
lcoords[0] = chartMetrics.get("canvasw").toInt()
}
}
else
{
lcoords = IntArray(6)
lcoords[0] = chartMetrics.get("canvasw").toInt()
}
val ldist = lcoords!![0] - chartMetrics.get("w")    // save distance between legend box and w (significant if legend is on rhs)
//io.starter.toolkit.Logger.log("Before Adjustments:  x:" + chartMetrics.get("x") + " w:" + chartMetrics.get("w") + " cw:" + chartMetrics.get("canvasw") + " y:" + chartMetrics.get("y") + " h:" + chartMetrics.get("h") + " ch:" + chartMetrics.get("canvash"));
// now adjust plot area coordinates based on canvas w, h, title and label offsets, and legend box, if any
chartMetrics.put("x", chartMetrics.get("x") + this.axes!!.axisMetrics.get("YAXISLABELOFFSET") as Double + this.axes!!.axisMetrics.get("YAXISTITLEOFFSET") as Double)
chartMetrics.put("y", chartMetrics.get("y") + chartMetrics.get("TITLEOFFSET"))
// TODO: seems that w is different doesn't need decrementing by x?? check out ...
chartMetrics.put("w", chartMetrics.get("w") - this.axes!!.axisMetrics.get("YAXISLABELOFFSET") as Double)
chartMetrics.put("h", chartMetrics.get("canvash") - chartMetrics.get("y") - this.axes!!.axisMetrics.get("XAXISLABELOFFSET") as Double - this.axes!!.axisMetrics.get("XAXISTITLEOFFSET") as Double - 10.0)
//io.starter.toolkit.Logger.log("After Adjustments:   x:" + chartMetrics.get("x") + " w:" + chartMetrics.get("w") + " cw:" + chartMetrics.get("canvasw") + " y:" + chartMetrics.get("y") + " h:" + chartMetrics.get("h") + " ch:" + chartMetrics.get("canvash"));

var cw = chartMetrics.get("canvasw")
// rhs legend has to have some extra adjustments to w and/or canvasw ...
if (lcoords!![5] == Legend.RIGHT)
{
val legendBeg = lcoords!![0] - (chartMetrics.get("w") + chartMetrics.get("x"))
val legendEnd = cw - (lcoords!![0].toDouble() + lcoords!![2].toDouble() + adjust)

if (legendBeg < 0 || legendEnd < 0)
{ // try to adjust
if (legendEnd < 0)
{
chartMetrics.put("canvasw", (lcoords!![0].toDouble() + lcoords!![2].toDouble() + adjust))
cw = (lcoords!![0].toDouble() + lcoords!![2].toDouble() + adjust)
}
//						if (legendBeg < 0)
//							chartMetrics.put("w",  lcoords[0]-10.0-chartMetrics.get("x"));
}
if (this.axes!!.hasAxis(ChartConstants.XAXIS) && ldist > 0)
// pie, donut, don't
// ensure distance between legend box and edge of plot area remains the same
chartMetrics.put("w", lcoords!![0].toDouble() - chartMetrics.get("x") - ldist)
//io.starter.toolkit.Logger.log("Adjusted LCoords:  " + Arrays.toString(lcoords));

}
else
{
val w = chartMetrics.get("w") + chartMetrics.get("x")
if (w > cw)
chartMetrics.put("w", chartMetrics.get("canvasw") - chartMetrics.get("x") - 10.0)
}

metricsDirty = false
}
catch (e:Exception) {
Logger.logErr("Chart.getMetrics: " + e.toString())
}

}
return chartMetrics
}

companion object {
internal val serialVersionUID = 6702247464633674375L

/**
 * generic method to find a specific record in the list of recs in chartarr
 *
 * @param chartArr
 * @param c        class of record to find
 * @return biffrec or null
*/
fun findRec(chartArr:ArrayList<*>, c:Class<*>):BiffRec? {
for (i in chartArr.indices)
{
val b = chartArr.get(i) as BiffRec
if (b.javaClass == c)
return b
}
return null
}

/**
 * generic method to find a specific record in the list of recs in chartArr
 *
 * @param c class of record to find
 * @return position of record
*/
fun findRecPosition(chartArr:ArrayList<*>, c:Class<*>):Int {
for (i in chartArr.indices)
{
val b = chartArr.get(i) as BiffRec
if (b.javaClass == c)
return i
}
return -1
}
}
}/**
 * returns the 3d record of the desired chart
 *
 * @param nChart 0=default, 1-9= overlay charts
 * @return
*/

