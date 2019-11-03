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
import io.starter.formats.OOXML.SpPr
import io.starter.formats.XLS.BiffRec
import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.Stack


/**
 * **ChartFormat: Parent Record for Chart Group (0x1014)**
 *
 *
 * 4 reserved 16 0 20 grbit 2 format flags 22 icrt 2 drawing order (0= bottom of
 * z-order)
 *
 *
 *
 *
 * 16 bytes- reserved must be 0
 * fVaried (1 bit): A bit that specifies whether the color for each data point and
 * the color and type for each data marker vary. If the chart group has multiple series or the chart group has one
 * series and the chart group type is a surface, stock, or area, this field MUST
 * be ignored, and the data points do not vary. For all other chart group types,
 * if the chart group has one series, a value of 0x1 specifies that the data
 * points vary.
 * 15 bits - reserved - 0
 * icrt (2 bytes): An unsigned integer that specifies the drawing order of the chart group relative to the other chart
 * groups, where 0x0000 is the bottom of the z-order.
 * This value MUST be unique for each instance of this record and MUST be less than or equal to 0x0009.
 *
 *
 *
 *
 * ORDER OF SUBRECS:
 * Bar/Pie/Scatter ...
 * ChartFormatLink
 * [SeriesList]
 * [ThreeD]
 * [Legend]
 * [DropBar]
 * [ChartLine, LineFormat]
 * [DataLabExt]
 * [DefaultText, Text]
 * [DataLabExtContents]
 * [DataFormat]
 * [ShapePropsStream]
 *
 *
 * "http://www.extentech.com">Extentech Inc.
 */
class ChartFormat : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    private var fVaried = false
    private var drawingOrder: Short = 0

    /**
     * returns whether
     * the color for each data point and the color and type for each data marker
     * vary
     *
     * @param b
     */
    /**
     * specifies whether
     * the color for each data point and the color and type for each data marker
     * vary
     *
     * @param b
     */
    var varyColor: Boolean
        get() = fVaried
        set(vary) {
            fVaried = vary
            grbit = ByteTools.updateGrBit(grbit, fVaried, 0)
            updateRecord()
        }

    /**
     * return truth of "chart is 3d clustered"
     *
     * @return
     */
    val is3DClustered: Boolean
        get() {
            val td = Chart.findRec(this.chartArr, ThreeD::class.java) as ThreeD
            return td?.isClustered ?: false
        }

    /**
     * returns an int representing the space between points in a 3d area, bar or line chart, or 0 if not 3d
     *
     * @return
     */
    /**
     * sets the Space between points (50 or 150 is default)
     *
     * @param gap
     */
    var gapDepth: Int
        get() {
            val td = getThreeDRec(false)
            td?.pcGap
            return 0
        }
        set(gap) {
            val td = getThreeDRec(true)
            td!!.pcGap = gap
        }

    /**
     * return ThreeD options in XML form
     *
     * @return String XML
     */
    val threeDXML: String
        get() {
            for (i in chartArr.indices) {
                val b = chartArr[i]
                if (b.opcode == XLSConstants.THREED)
                    return (b as ThreeD).optionsXML
            }
            return ""
        }


    /**
     * return the Data Labels chosen for this chart, if any can be one or more
     * of: <br></br>
     * Value <br></br>
     * ValuePerecentage <br></br>
     * CategoryPercentage <br></br>
     * CategoryLabel <br></br>
     * BubbleLabel <br></br>
     * SeriesLabel or an empty string if no data labels are chosen for the chart
     *
     * @return
     */
    val dataLabels: String?
        get() {
            val df = this.getDataFormatRec(false)
            return if (df != null) {
                df.dataLabelType
            } else ""
        }

    /**
     * returns true if this chart displays lines (Line, Scatter, Radar)
     *
     * @return true if chart has lines (see Scatter, Line chart ...)
     */
    val hasLines: Boolean
        get() {
            val df = this.getDataFormatRec(false)
            return df?.hasLines ?: false
        }

    /**
     * returns true if this chart has smoothed lines (Scatter, Line, Radar)
     *
     * @return
     */
    /**
     * sets this chart to have smoothed lines (Scatter, Line, Radar)
     */
    var hasSmoothLines: Boolean
        get() {
            val df = this.getDataFormatRec(false)
            return df?.smoothedLines ?: false
        }
        set(b) {
            val df = this.getDataFormatRec(true)
            df!!.setSmoothLines(b)
        }

    /**
     * returns true if chart has Drop Lines
     *
     * @return
     */
    /* chartline:
		line, chartformatlink, <serieslist>, <3d>, <legend>, chartline, lineformat, startblock,
		shapepropsstream, [-92, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
		                   0, 0,
		                   -51, -110,
		                   -66, 74, 1, -88, 0, 0, 0, 0]
		endblock

	dropbar= chartline, lineformat
	*/ val hasDropLines: Boolean
        get() = false

    /**
     * return data label options as an int <br></br>
     * can be one or more of: <br></br>
     * SHOWVALUE= 0x1; <br></br>
     * SHOWVALUEPERCENT= 0x2; <br></br>
     * SHOWCATEGORYPERCENT= 0x4; <br></br>
     * SHOWCATEGORYLABEL= 0x10; <br></br>
     * SHOWBUBBLELABEL= 0x20; <br></br>
     * SHOWSERIESLABEL= 0x40;
     *
     * @return a combination of data label options above or 0 if none
     *
     *
     * NOTE: this returns the Data Labels settings for the entire chart,
     * not a particular series
     * @see AttachedLabel
     */
    // here we are assuming that the TextDisp is of the proper ObjectLink=4 type ...
    // Extended Label -- add to attachedlabel, if any
    // if so, no fontx record ... use default???
    val dataLabelsInt: Int
        get() {
            var datalabels = 0
            val z = Chart.findRecPosition(this.chartArr, DataLabExtContents::class.java)
            if (z > 0 && this.chartArr[z - 1] is TextDisp) {
                val dl = chartArr[z] as DataLabExtContents
                datalabels = dl.typeInt
            }
            val df = this.getDataFormatRec(false)
            if (df != null) {
                datalabels = datalabels or df.dataLabelTypeInt
            }
            return datalabels
        }

    /**
     * returns the bar shape for a column or bar type chart can be one of: <br></br>
     * ChartConstants.SHAPECOLUMN default <br></br>
     * ChartConstants.SHAPECONE <br></br>
     * ChartConstants.SHAPECONETOMAX <br></br>
     * ChartConstants.SHAPECYLINDER <br></br>
     * ChartConstants.SHAPEPYRAMID <br></br>
     * ChartConstants.SHAPEPYRAMIDTOMAX
     *
     * @return
     */
    /**
     * bar shape:
     * <br></br>the shape is as follows:
     * public static final int SHAPECOLUMN= 0;		// default
     * public static final int SHAPECYLINDER= 1;
     * public static final int SHAPEPYRAMID= 256;
     * public static final int SHAPECONE= 257;
     * public static final int SHAPEPYRAMIDTOMAX= 516;
     * public static final int SHAPECONETOMAX= 517;
     *
     * @param shape
     */
    // THIS DOES NOT MAKE SENSE ACCORDING TO DOC BUT IS WHAT EXCEL DOES
    var barShape: Int
        get() {
            var shape = ChartConstants.SHAPEDEFAULT
            val df = this.getDataFormatRec(false)
            if (df != null)
                shape = df.shape.toInt()
            return shape
        }
        set(shape) {
            val df = this.getDataFormatRec(true)
            df!!.setPointNumber(0)
            df.setSeriesIndex(0)
            df.setSeriesNumber(-3)
            df.setShape(shape)
        }

    /**
     * returns type of marker, if any <br></br>
     * 0 = no marker <br></br>
     * 1 = square <br></br>
     * 2 = diamond <br></br>
     * 3 = triangle <br></br>
     * 4 = X <br></br>
     * 5 = star <br></br>
     * 6 = Dow-Jones <br></br>
     * 7 = standard deviation <br></br>
     * 8 = circle <br></br>
     * 9 = plus sign
     *
     * @return
     */
    // default= circles
    // default actually looks like: 2, 1, 5, 4 ...
    val markerFormat: Int
        get() {
            val markertype = 8
            val df = this.getDataFormatRec(false)
            return df?.markerFormat ?: 0
        }

    /**
     * return ChartLines option, if any
     * <pre>
     * 0= drop lines below the data points of Line, Area and Stock charts
     * 1= High-low lines around the data points of Line and Stock charts
     * 2- Series Line connecting data points of stacked column and bar charts and OfPie Charts
    </pre> *
     *
     * @return 0-2 or -1 if no chart lines
     */
    val chartLines: Int
        get() {
            val cl = Chart.findRec(chartArr, ChartLine::class.java) as ChartLine
            return cl?.lineType ?: -1
        }

    /**
     * return the record governing chart lines: dropLines, Hi-low lines or Series Lines
     *
     * @return
     */
    val chartLinesRec: ChartLine
        get() = Chart.findRec(chartArr, ChartLine::class.java) as ChartLine

    /**
     * return the OOXML to define the upDownBars element
     * <br></br>defined by 2 Dropbar records in this subarray
     * <br></br>Only valid for Line and Stock charts
     *
     * @return
     */
    //
    // c:gapWidth
    // TODO: only upBar?  Should they match?
    // default
    // c:upBars
    // c:downBars
    val upDownBarOOXML: String
        get() {
            var z = Chart.findRecPosition(chartArr, Dropbar::class.java)
            if (z == -1) return ""
            val ooxml = StringBuffer()
            try {
                val upBars = chartArr[z++] as Dropbar
                val downBars = chartArr[z] as Dropbar
                z++
                ooxml.append("<c:upDownBars>")
                val gw = upBars.gapWidth.toInt()
                if (gw != 150)
                    ooxml.append("<c:gapWidth val=\"$gw\"/>")
                ooxml.append(upBars.getOOXML(true))
                ooxml.append(downBars.getOOXML(false))
                ooxml.append("</c:upDownBars>")
            } catch (e: Exception) {
            }

            return ooxml.toString()
        }

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    /**
     * access chart-type record and return any options specific for this chart
     * in XML form Gathers chart options such as show legend, grid lines, etc
     * ... these options are distinct from chart-type-specific options, which
     * are handled by the appropriate chart record (pie, bar ...) Also, since
     * both ThreeD options and Axis-specific options are quite extensive, these
     * are handled separately
     *
     * @return String of options XML ("" if no options set)
     * @see setChartOption
     *
     * @see setOption
     *
     * @see getThreeDXML
     * @see getAxesXML
     */
    // controls Data Legends, % Distance of sections, Line
    // Format, Area Format, Bar Shapes ...
    // controls Show Legend + // some Data Legends
    val chartOptionsXML: String
        get() {
            var bFoundDefaultText0 = false
            val sb = StringBuffer()
            for (i in chartArr.indices) {
                var b: BiffRec = chartArr[i]
                if (b is ChartObject) {
                    val co = b as ChartObject
                    if (b is DataFormat) {
                        val df = b
                        val shape = df.shape.toInt()
                        if (shape != 0)
                            sb.append(" Shape=\"$shape\"")
                        for (z in df.chartArr.indices) {
                            b = df.chartArr[z]
                            if (b is PieFormat) {
                                sb.append(b.optionsXML)
                            } else if (b is AttachedLabel) {
                                val type = b.type
                                sb.append(" DataLabel=\"$type\"")
                            } else if (b is Serfmt) {
                                sb.append(b.optionsXML)
                            } else if (b is MarkerFormat) {
                                sb.append(b.optionsXML)
                            }
                        }
                    } else if (b is DefaultText) {
                        if (!bFoundDefaultText0)
                            bFoundDefaultText0 = b.type == 0
                        if (b.type == 1 && bFoundDefaultText0)
                            sb.append(" ShowLegendKey=\"true\"")
                    }
                }
            }
            return sb.toString()
        }

    override fun init() {
        super.init()
        grbit = ByteTools.readShort(this.getByteAt(16).toInt(), this.getByteAt(17).toInt())
        drawingOrder = ByteTools.readShort(this.getByteAt(18).toInt(),
                this.getByteAt(19).toInt())
        fVaried = grbit and 0x1 == 0x1
    }

    /**
     *
     */
    private fun updateRecord() {
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[16] = b[0]
        this.data[17] = b[1]
    }

    /**
     * replace the existing chart object with the desired ChartObject,
     * effectively changing the type of the chart
     *
     * @param co
     */
    fun setChartObject(co: ChartObject) {
        chartArr.removeAt(0)
        chartArr.add(0, co as XLSRecord)

    }

    /**
     * @return truth of "Chart is Three D"
     */
    fun isThreeD(chartType: Int): Boolean {
        if (chartType != ChartConstants.BUBBLECHART) {
            for (i in chartArr.indices) {
                val b = chartArr[i]
                if (b.opcode == XLSConstants.THREED)
                    return true
            }
        } else {
            val df = this.getDataFormatRec(false)
            if (df != null) {
                return df.has3DBubbles
            }
        }
        return false
    }

    /**
     * sets if this chart has clustered bar/columns
     *
     * @param bIsClustered
     */
    fun setIsClustered(bIsClustered: Boolean) {
        val td = getThreeDRec(false)
        if (td != null) {
            td.isClustered = bIsClustered
        } else {
            if (chartArr[0].opcode == XLSConstants.BAR) {
                (chartArr[0] as Bar).setIsClustered()
            }
        }

    }

    /**
     * percentage=distance of pie slice from center of pie as %
     *
     * @param p
     */
    fun setPercentage(p: Int) {
        val df = this.getDataFormatRec(true)
        df!!.percentage = p
    }

    /**
     * sets chart options such as threed options, grid lines, etc ... these options
     * are distinct from chart-type-specific options, which are handled by the
     * appropriate chart record (pie, bar ...)
     *
     * @param op  String option name
     * @param val Object value
     */
    fun setOption(op: String, `val`: String) {
        if (op.equals("Percentage", ignoreCase = true)) {
            setPercentage(java.lang.Short.valueOf(`val`).toShort().toInt())
        } else if (op.equals("Shape", ignoreCase = true)) {
            barShape = Integer.parseInt(`val`)
        } else if (op == "ShowBubbleSizes"
                || // TextDisp options

                op == "ShowLabelPct" || op == "ShowCatLabel"
                || op == "ShowPct" || op == "Rotation" ||
                // op.equals("ShowValue") || unknown
                op == "Label" || op == "TextRotation") {
            val td = getDataLegendTextDisp(0)
            td!!.setChartOption(op, `val`)
        } else if (op == "Perspective"
                || // ThreeD options

                op == "Cluster" || op == "ThreeDScaling"
                || op == "TwoDWalls" || op == "PcGap"
                || op == "PcDepth" || op == "PcHeight"
                || op == "PcDist" || op == "AnElev"
                || op == "AnRot") {
            val td = this.getThreeDRec(true)
            td!!.setChartOption(op, `val`)
        } else if (op == "ShowValueLabel"
                || // Attached Label Options

                op == "ShowValueAsPercent"
                || op == "ShowLabelAsPercent"
                || op == "ShowLabel"
                || op == "ShowSeriesName" || op == "ShowBubbleLabel") {
            val df = this.getDataFormatRec(true)
            df!!.setDataLabels(op)
        } else if (op.equals("SmoothedLine", ignoreCase = true)
                || op.equals("ThreeDBubbles", ignoreCase = true)
                || op.equals("ArShadow", ignoreCase = true)) {
            val df = this.getDataFormatRec(true)
            if (op.equals("SmoothedLine", ignoreCase = true))
                df!!.setSmoothLines(true)
            else if (op.equals("ThreeDBubbles", ignoreCase = true))
                df!!.has3DBubbles = true
            else
                df!!.hasShadow = true
        }
    }

    /**
     * Return the ThreeD rec associated with this ChartFormat, create if not
     * present
     */
    // NOTES:
    // The Chart3d record specifies that the plot area of the chart group is rendered in a 3-D scene
    // and also specifies the attributes of the 3-D plot area.
    // The preceding chart group type MUST be of type bar, pie, line, area, or surface.
    fun getThreeDRec(bCreate: Boolean): ThreeD? {
        var td = Chart.findRec(this.chartArr, ThreeD::class.java) as ThreeD
        if (td == null && bCreate) { // add ThreeD rec
            var i = 0
            while (i < this.chartArr.size) {
                val b = this.chartArr[i]
                if (b.opcode == XLSConstants.CHARTFORMATLINK) {
                    if (i + 1 < chartArr.size && (this.chartArr[i + 1] as BiffRec).opcode == XLSConstants.SERIESLIST)
                        i++    // rare that SeriesList record appears
                    td = ThreeD.prototype as ThreeD?
                    td.parentChart = this.parentChart
                    this.chartArr.add(i + 1, td)
                    break
                }
                i++
            }
        }
        return td
    }

    /**
     * Add or Retrieve TextDisp and assoc records specific for Data Legends
     *
     * @return
     */
    private fun getDataLegendTextDisp(type: Int): TextDisp? {
        var i = Chart.findRecPosition(this.chartArr, Legend::class.java)
        var td: TextDisp? = null
        if (this.chartArr.size <= i + 1 || this.chartArr[i + 1].javaClass != DefaultText::class.java) { // then add one
            val d = DefaultText.prototype as DefaultText?
            d!!.setType(type.toShort())
            d.parentChart = this.parentChart
            this.chartArr.add(++i, d)
            td = TextDisp.getPrototype(ObjectLink.TYPE_DATAPOINTS,
                    "", this.workBook) as TextDisp
            td.parentChart = this.parentChart
            this.chartArr.add(++i, td)
        } else {
            var d = this.chartArr[i + 1] as DefaultText
            if (d.type != type) { // / add a new one
                i += 2 // add after TextDisp
                d = DefaultText.prototype as DefaultText?
                d.setType(type.toShort())
                d.parentChart = this.parentChart
                this.chartArr.add(++i, d)
                td = TextDisp.getPrototype(
                        ObjectLink.TYPE_DATAPOINTS, "", this.workBook) as TextDisp
                td.parentChart = this.parentChart
                this.chartArr.add(++i, td)
            } else { // it's the correct one
                i += 2
                td = this.chartArr[i] as TextDisp
            }
        }
        return td
    }

    /**
     * Gets the dataformat record associated with this ChartFormat If none
     * present, creates a basic DataFormat set of records DataFormat controls
     * Data Labels, % Distance from sections, line formats ...
     *
     * @return DataFormat Record
     */
    private fun getDataFormatRec(bCreate: Boolean): DataFormat? {
        var df = Chart.findRec(this.chartArr, DataFormat::class.java) as DataFormat
        if (df == null && bCreate) { // create dataformat
            df = DataFormat.getPrototypeWithFormatRecs(this.parentChart) as DataFormat
            this.addChartRecord(df)
        }
        return df
    }


    /**
     * sets this chart to have default lines (Scatter, Line chart ...)
     */
    fun setHasLines() {
        val df = this.getDataFormatRec(true)
        df!!.setHasLines()
    }


    /**
     * sets this chart to have lines (Scatter, Line chart ...) of style lineStyle
     * <br></br>Style of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
     */
    fun setHasLines(lineStyle: Int) {
        val df = this.getDataFormatRec(true)
        df!!.setHasLines(lineStyle)
    }

    /**
     * sets this chart (Area, Line or Stock) to have Drop Lines
     */
    fun setHasDropLines() {

    }

    /**
     * sets 3d bubble state
     *
     * @param has3dBubbles
     */
    fun setHas3DBubbles(has3dBubbles: Boolean) {
        val df = this.getDataFormatRec(true)
        df!!.has3DBubbles = has3dBubbles

    }

    /**
     * sets the data labels for the entire chart (as opposed to a specific series/data point).
     * A combination of:
     *  * SHOWVALUE= 0x1;
     *  * SHOWVALUEPERCENT= 0x2;
     *  * SHOWCATEGORYPERCENT= 0x4;
     *  * SHOWCATEGORYLABEL= 0x10;
     *  * SHOWBUBBLELABEL= 0x20;
     *  * SHOWSERIESLABEL= 0x40;
     */
    fun setHasDataLabels(dl: Int) {
        val df = this.getDataFormatRec(true)
        df!!.setHasDataLabels(dl)
    }

    /**
     * 0 = no marker <br></br>
     * 1 = square <br></br>
     * 2 = diamond <br></br>
     * 3 = triangle <br></br>
     * 4 = X <br></br>
     * 5 = star <br></br>
     * 6 = Dow-Jones <br></br>
     * 7 = standard deviation <br></br>
     * 8 = circle <br></br>
     * 9 = plus sign
     */
    fun setMarkers(markerFormat: Int) {
        val df = this.getDataFormatRec(true)
        df!!.markerFormat = markerFormat
    }

    /**
     * Chart Line options (not available on all charts):
     * <br></br>Drop Lines available on Line, Area and Stock charts
     * <br></br>HiLow Lines are available on Line and Stock charts
     * <br></br>Series Lines are available on Bar and OfPie charts
     * <pre>
     * 0= drop lines below the data points of Line, Area and Stock charts
     * 1= High-low lines around the data points of Line and Stock charts
     * 2- Series Line connecting data points of stacked column and bar charts and OfPie Charts
    </pre> *
     * <br></br>
     *
     * @param lineType
     */
    fun addChartLines(lineType: Int): ChartLine {
        //		ChartLine cl= (ChartLine) Chart.findRec(chartArr, ChartLine.class);
        //		if (cl==null) {
        var cl: ChartLine? = null
        // goes After DropBar or legend or 3d (or CFL)
        var i = chartArr.size - 1
        while (i >= 1) {
            val br = chartArr[i]
            val op = br.opcode
            if (op == XLSConstants.DROPBAR || op == XLSConstants.LEGEND || op == XLSConstants.THREED || op == XLSConstants.CHARTFORMATLINK) {
                cl = ChartLine.prototype as ChartLine?
                cl!!.parentChart = this.parentChart
                chartArr.add(++i, cl)
                cl.parentChart = this.parentChart    // ensure can find
                val lf = LineFormat.getPrototype(0, 1) as LineFormat
                lf.parentChart = this.parentChart
                chartArr.add(++i, lf)
                break
            }
            i--
        }
        //		}
        cl!!.lineType = lineType
        return cl
    }

    fun getChartLinesRec(id: Int): ChartLine? {
        var i = Chart.findRecPosition(chartArr, ChartLine::class.java)
        if (i > -1) {
            while (i < chartArr.size) {
                if ((chartArr[i] as BiffRec).opcode == XLSConstants.CHARTLINE) {
                    val cl = chartArr[i] as ChartLine
                    if (cl.lineType == id)
                        return cl
                } else
                    break
                i += 2
            }
        }
        return null
    }

    /**
     * add up/down bars (line, area stock)
     */
    fun addUpDownBars() {
        if (Chart.findRec(chartArr, Dropbar::class.java) == null) {
            // create necessary records to describe up/down bars
            val upBar = Dropbar.prototype as Dropbar?
            upBar!!.parentChart = this.parentChart
            val downBar = Dropbar.prototype as Dropbar?
            downBar!!.parentChart = this.parentChart
            var lf: LineFormat = LineFormat.prototype as LineFormat?
            lf.parentChart = this.parentChart
            var af: AreaFormat = AreaFormat.prototype as AreaFormat?
            af.parentChart = this.parentChart
            upBar.chartArr.add(lf)
            upBar.chartArr.add(af)
            lf = LineFormat.prototype as LineFormat?
            lf.parentChart = this.parentChart
            af = AreaFormat.prototype as AreaFormat?
            af.parentChart = this.parentChart
            downBar.chartArr.add(lf)
            downBar.chartArr.add(af)

            // add dropbar records to subarray
            var i = chartArr.size - 1
            while (i >= 0) {
                val br = chartArr[i]
                val op = br.opcode
                if (op == XLSConstants.SERIESLIST || op == XLSConstants.LEGEND || op == XLSConstants.THREED || op == XLSConstants.CHARTFORMATLINK) {
                    chartArr.add(++i, upBar)
                    chartArr.add(++i, downBar)
                    break
                }
                i--
            }
        }
    }


    /**
     * parse upDownBars OOXML element (controled by 2 DropBar records in this subArray)
     * <br></br>Valid for Line and Stock charts only
     */
    fun parseUpDownBarsOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle) {
        // assume dropbar records are not present ...
        addUpDownBars()
        var z = Chart.findRecPosition(chartArr, Dropbar::class.java)
        val downBar = chartArr[z++] as Dropbar
        val upBar = chartArr[z] as Dropbar


        try {
            var curbar: Dropbar? = null
            var eventType = xpp.eventType
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    val tnm = xpp.name
                    if (tnm == "downBars") {
                        curbar = downBar
                    } else if (tnm == "upBars") {
                        curbar = upBar
                    } else if (tnm == "gapWidth") {    // default=150
                        upBar.setGapWidth(Integer.valueOf(xpp.getAttributeValue(0)))    // TODO: should this be in 1st dropbar or both??
                        downBar.setGapWidth(Integer.valueOf(xpp.getAttributeValue(0)))    // TODO: should this be in 1st dropbar or both??
                    } else if (tnm == "spPr") {
                        lastTag.push(tnm)
                        val sppr = SpPr.parseOOXML(xpp, lastTag, bk).cloneElement() as SpPr
                        val lf = curbar!!.chartArr[0] as LineFormat
                        lf?.setFromOOXML(sppr)
                        // TODO: fill AreaFormat with sppr
                    }
                } else if (eventType == XmlPullParser.END_TAG) {
                    if (xpp.name == "upDownBars") {
                        lastTag.pop()
                        break
                    }
                }
                eventType = xpp.next()
            }
        } catch (e: Exception) {
            Logger.logErr("parseUpDownBarsOOXML: $e")
        }

    }


    /**
     * return the drawing order of this ChartFormat <br></br>
     * For multiple charts-in-one, drawing order determines the order of the
     * charts
     *
     * @return
     */
    fun getDrawingOrder(): Int {
        return drawingOrder.toInt()
    }

    /**
     * set the drawing order of this ChartFormat <br></br>
     * For multiple charts-in-one, drawing order determines the order of the
     * charts
     *
     * @param order
     */
    fun setDrawingOrder(order: Int) {
        drawingOrder = order.toShort()
        val b = ByteTools.shortToLEBytes(drawingOrder)
        this.data[18] = b[0]
        this.data[19] = b[1]
    }

    /**
     * get the value of *almost* any chart option (axis options are in Axis)
     *
     * @param op String option e.g. Shadow or Percentage
     * @return String value of option
     */
    override fun getChartOption(op: String): String? {
        val df = this.getDataFormatRec(false)
        try {
            if (op == "Percentage") { // pieformat
                return df!!.percentage.toString()
            } else if (op == "ShowValueLabel"
                    || // Attached Label Options

                    op == "ShowValueAsPercent"
                    || op == "ShowLabelAsPercent"
                    || op == "ShowLabel"
                    || op == "ShowBubbleLabel") {
                return df!!.getDataLabelType(op)
            } else if (op == "ShowBubbleSizes"
                    || // TextDisp options

                    op == "ShowLabelPct" || op == "ShowPct"
                    || op == "ShowCatLabel"
                    ||
                    // op.equals("ShowValue") || unknown
                    op == "Rotation" || op == "Label"
                    || op == "TextRotation") {
                val td = getDataLegendTextDisp(0)
                return td!!.getChartOption(op)
            } else if (op == "Perspective"
                    || // ThreeD options

                    op == "Cluster" || op == "ThreeDScaling"
                    || op == "TwoDWalls" || op == "PcGap"
                    || op == "PcDepth" || op == "PcHeight"
                    || op == "PcDist" || op == "AnElev"
                    || op == "AnRot") {
                val td = this.getThreeDRec(false)
                return if (td != null) td.getChartOption(op) else ""
            } else if (op == "ThreeDBubbles") {
                return df!!.has3DBubbles.toString()
            } else if (op == "ArShadow") {
                return df!!.hasShadow.toString()
            } else if (op == "SmoothLines") {
                return df!!.smoothedLines.toString()
                // TODO: FINSIH REST!
            } else if (op == "AxisLabels") { // Radar, RadarArea
            } else if (op == "BubbleSizeRatio") { // Scatter
            } else if (op == "BubbleSize") { // Scatter
            } else if (op == "ShowNeg") { // Scatter
            } else if (op == "ColorFill") { // Surface
            } else if (op == "Shading") { // Surface
            } else if (op == "MarkerFormat") { // MarkerFormat
            }
        } catch (e: NullPointerException) {

        }

        return ""
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 4000704166442059677L

        val prototype: XLSRecord?
            get() {
                val cf = ChartFormat()
                cf.opcode = XLSConstants.CHARTFORMAT
                cf.data = cf.PROTOTYPE_BYTES
                cf.init()
                return cf
            }
    }

}
