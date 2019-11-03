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


import io.starter.OpenXLS.CellHandle
import io.starter.OpenXLS.CellRange
import io.starter.OpenXLS.FormatHandle
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.OOXML.DLbls
import io.starter.formats.OOXML.DPt
import io.starter.formats.OOXML.Marker
import io.starter.formats.OOXML.SpPr
import io.starter.formats.XLS.*
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.ArrayList
import java.util.Vector

/**
 * **Series: Series Definition (1003h)**<br></br>
 *
 *
 * This record defines the Series data of a chart.
 *
 *
 * sdtX and sdtY fields determine data type (numeric and text)
 * cValx and cValy fields determine number of cells in series
 *
 *
 * Offset           Name    Size    Contents
 * --
 * 4               sdtX    2       Type of data in cats (1=num, 3=str)
 * 8               sdtY    2       Type of data in values (1=num, 3=str)
 * 10              cValx   2       Count of categories
 * 12              cValy	2       Count of Values
 * 14              sdtBSz  2		Type of data in Bubble size series (0=dates, 1=num, 2=seq., 3=text)
 * 16           sdtValBSz  2		Count of Bubble series vals
 *
 *
 * sdtX (2 bytes): An unsigned integer that specifies the type of data in categories (3), or horizontal values on bubble and scatter chart groups, in the series.
 * value :0x0001		The series contains categories), or horizontal values on bubble and scatter chart groups, with numeric information.
 * value:  0x0003		The series contains categories, or horizontal values on bubble and scatter chart groups, with text information.
 * sdtY (2 bytes): An unsigned integer that specifies that the values, or vertical values on bubble and scatter chart groups, in the series contain numeric information.
 * MUST be set to 0x0001, and MUST be ignored.
 * cValx (2 bytes): An unsigned integer that specifies the count of categories (3), or horizontal values on bubble and scatter chart groups, in the series.
 * This value MUST be less than or equal to 0x0F9F.
 * cValy (2 bytes): An unsigned integer that specifies the count of values, or vertical values on bubble and scatter chart groups, in the series.
 * This value MUST be less than or equal to 0x0F9F.
 * sdtBSize (2 bytes): An unsigned integer that specifies that the bubble size values in the series contain numeric information.
 * This value MUST be set to 0x0001, and MUST be ignored.
 * cValBSize (2 bytes): An unsigned integer that specifies the count of bubble size values in the series.
 * This value MUST be less than or equal to 0x0F9F.
 *
 *
 *
 *
 *
 *
 * The series object contains a collection of sub objects.  Usually this will take the form of 4 ai records,
 * type 0-3, and supporting records such as labels.
 *
 *
 *
 * @see Chart
 */

/**
 * sdtX (2 bytes): An unsigned integer that specifies the type of data in categories (3), or horizontal values on bubble and scatter chart groups, in the series. MUST be a value from the following table.
 * Value 	Meaning
 * 0x0001	The series contains categories, or horizontal values on bubble and scatter chart groups, with numeric information.
 * 0x0003	The series contains categories, or horizontal values on bubble and scatter chart groups, with text information.
 * sdtY (2 bytes): An unsigned integer that specifies that the values, or vertical values on bubble and scatter chart groups, in the series contain numeric information. MUST be set to 0x0001, and MUST be ignored.
 * cValx (2 bytes): An unsigned integer that specifies the count of categories (3), or horizontal values on bubble and scatter chart groups, in the series. This value MUST be less than or equal to 0x0F9F.
 * cValy (2 bytes): An unsigned integer that specifies the count of values, or vertical values on bubble and scatter chart groups, in the series. This value MUST be less than or equal to 0x0F9F.
 * sdtBSize (2 bytes): An unsigned integer that specifies that the bubble size values in the series contain numeric information. This value MUST be set to 0x0001, and MUST be ignored.
 * cValBSize (2 bytes): An unsigned integer that specifies the count of bubble size values in the series. This value MUST be less than or equal to 0x0F9F.
 */
class Series : GenericChartObject(), ChartObject {

    protected var sdtX = -1
    protected var sdtY = -1
    protected var cValx = -1
    protected var cValy = -1
    protected var sdtBSz = -1
    protected var sdtValBSz = -1

    /**
     * return the OOXML shape property for this series
     *
     * @return
     */
    /**
     * set the OOXML shape properties for this series
     *
     * @param sp
     */
    var spPr: SpPr? = null    // OOXML-specific holds the shape properties (line and fill) for this series (all charts)
    /**
     * return the OOXML marker properties for this series
     *
     * @return
     */
    /**
     * set the OOXML marker properties for this series
     *
     * @param Sp
     */
    var marker: Marker? = null            // OOXML-specific object to hold marker properties for this series (radar, scatter and line charts only)
    /**
     * return the OOXML dLbls (data labels) properties for this series
     *
     * @return
     */
    /**
     * set the OOXML dLbls (data labels) properties for this series
     *
     * @param Sp
     */
    var dLbls: DLbls? = null            // OOXML-specific object holds Data Labels properties for this series (all charts except surface)
    private var dPts: ArrayList<*>? = null    // OOXML-specific object holds Data Labels properties for this series (all charts except surface)

    /**
     * Returns the series value AI associated with this series.
     *
     * @return
     */
    val seriesValueAi: Ai?
        get() {
            for (i in chartArr.indices) {
                val br = chartArr[i]
                if (br.opcode == XLSConstants.AI) {
                    val thisAi = br as Ai
                    if (thisAi.type == Ai.TYPE_VALS)
                        return thisAi
                }
            }
            return null
        }

    /**
     * returns the custom number format for this series, 0 if none
     *
     * @return
     */
    private// if 0, number format is determined by the application
    // meaning it uses the number format of the source data
    val seriesNumberFormat: Int
        get() {
            val ai = this.seriesValueAi
            var i = 0
            if (ai != null) {
                i = ai.ifmt
                if (i != 0)
                    return i
                try {
                    val p = ai.cellRangePtgs[0].components[0] as io.starter.formats.XLS.formulas.PtgRef
                    i = ai.workBook!!.getXf(p.refCells!![0].ixfe)!!.ifmt.toInt()
                } catch (e: Exception) {
                }

            }
            return i
        }

    /**
     * return the String representation of the numeric format pattern for the series (values) axis
     *
     * @return
     */
    // custom??
    val seriesFormatPattern: String?
        get() {
            val ifmt = seriesNumberFormat
            val fmts = FormatConstantsImpl.builtinFormats
            for (x in fmts.indices) {
                if (ifmt == Integer.parseInt(fmts[x][1], 16))
                    return fmts[x][0]
            }
            try {
                val fmt = this.workBook!!.getFormat(ifmt)
                return fmt.format
            } catch (e: Exception) {
            }

            return "General"
        }

    /**
     * returns the custom number format for a value-type axis
     *
     * @return
     */
    private val categoryNumberFormat: Int
        get() {
            val ai = this.categoryValueAi
            return ai?.ifmt ?: 0
        }

    /**
     * return the String representation of the numeric format pattern for the Catgeory axis
     *
     * @return
     */
    // custom??
    val categoryFormatPattern: String?
        get() {
            val ifmt = categoryNumberFormat
            val fmts = FormatConstantsImpl.builtinFormats
            for (x in fmts.indices) {
                if (ifmt == Integer.parseInt(fmts[x][1], 16))
                    return fmts[x][0]
            }
            try {
                val fmt = this.workBook!!.getFormat(ifmt)
                return fmt.format
            } catch (e: Exception) {
            }

            return "General"
        }


    /**
     * get legend text
     *
     * @return
     */
    // couldn't find it!
    val legendText: String
        get() {
            for (i in chartArr.indices) {
                val br = chartArr[i]
                if (br.opcode == XLSConstants.AI) {
                    val thisAi = br as Ai
                    if (thisAi.type == Ai.TYPE_TEXT) {
                        try {
                            val p = thisAi.cellRangePtgs
                            return (p[0] as io.starter.formats.XLS.formulas.PtgRef).formattedValue
                        } catch (e: Exception) {
                        }

                        try {
                            if (chartArr.size > i + 1) {
                                val st = chartArr[i + 1] as SeriesText
                                if (st != null)
                                    return st.toString()
                            }
                        } catch (e: ClassCastException) {
                        }

                    }
                }
            }
            return ""
        }

    /**
     * return the legend cell reference
     *
     * @return
     */
    /**
     * set legend to a cell ref.
     *
     * @param newLegendCell
     */
    //CellHandle cell= this.getWorkBook().getCell(newLegendCell);
    //    		newLegendCell= newLegendCell.replace('!', ':');	// for this method it's Sheet:cell (????)
    var legendRef: String?
        get() {
            val ai = this.legendAi
            return ai?.definition
        }
        set(newLegendCell) {
            val ai = this.legendAi
            ai!!.changeAiLocation(ai.toString(), newLegendCell)
            val st = this.legendSeriesText
            ai.setRt(2)
            var legendText = ""
            try {
                val r = ai.workBook!!.getCell(newLegendCell)
                legendText = r.stringVal
            } catch (e: Exception) {
                Logger.logErr("Series.setLegendRef: Error setting Legend Reference to '$newLegendCell': $e")
            }

            st!!.setText(legendText)
        }

    /**
     * Return the SeriesText object related to the Legend
     *
     * @return
     */
    protected// couldn't find it!
    val legendSeriesText: SeriesText?
        get() {
            for (i in chartArr.indices) {
                val br = chartArr[i]
                if (br.opcode == XLSConstants.AI) {
                    val thisAi = br as Ai
                    if (thisAi.type == Ai.TYPE_TEXT) {
                        if (chartArr.size > i + 1) {
                            try {
                                return chartArr[i + 1] as SeriesText
                            } catch (e: ClassCastException) {
                                return null
                            }

                        }
                    }
                }
            }
            return null
        }

    /**
     * Returns the legend value Ai associated with this series
     *
     * @return
     */
    val legendAi: Ai?
        get() {
            for (i in chartArr.indices) {
                val br = chartArr[i]
                if (br.opcode == XLSConstants.AI) {
                    val thisAi = br as Ai
                    if (thisAi.type == Ai.TYPE_TEXT)
                        return thisAi
                }
            }
            return null
        }


    //    	if (hasBubbleSizes()) {
    //    	}
    val bubbleValueAi: Ai?
        get() {
            for (i in chartArr.indices) {
                val br = chartArr[i]
                if (br.opcode == XLSConstants.AI) {
                    val thisAi = br as Ai
                    if (thisAi.type == Ai.TYPE_BUBBLES)
                        return thisAi
                }
            }
            return null
        }

    private val PROTOTYPE_BYTES = byteArrayOf(3, 0, 1, 0, 3, 0, 3, 0, 1, 0, 0, 0)

    /**
     * Returns the category value AI associated with this series.
     *
     * @return
     */
    val categoryValueAi: Ai?
        get() {
            for (i in chartArr.indices) {
                val br = chartArr[i]
                if (br.opcode == XLSConstants.AI) {
                    val thisAi = br as Ai
                    if (thisAi.type == Ai.TYPE_CATEGORIES)
                        return thisAi
                }
            }
            return null
        }


    /**
     * Get the series index (file relative)
     *
     * @return
     */
    protected val seriesIndex: Int
        get() {
            val df = this.getDataFormatRec(false)
            return df?.seriesIndex?.toInt() ?: -1
        }

    /**
     * Get the series Number (display)
     *
     * @return
     */
    protected val seriesNumber: Int
        get() {
            val df = this.getDataFormatRec(false)
            return df?.seriesNumber?.toInt() ?: -1
        }

    protected//20070711 KSC: changed from protected
    var categoryCount: Int
        get() = cValx
        set(i) {
            cValx = i
            this.update()
        }

    protected//20070711 KSC: changed from protected
    var valueCount: Int
        get() = cValy
        set(i) {
            cValy = i
            this.update()
        }

    protected//20070711 KSC: changed from protected
    var bubbleCount: Int
        get() = sdtValBSz
        set(i) {
            sdtValBSz = i
            this.update()
        }

    // 20070712 KSC: get/set for data types
    var categoryDataType: Int
        get() = sdtX
        set(i) {
            sdtX = i
            this.update()
        }

    var valueDataType: Int
        get() = sdtY
        set(i) {
            sdtY = i
            this.update()
        }

    /**
     * returns the shape of the data point for this series
     *
     * @return
     */
    var shape: Int
        get() {
            var ret = 0
            val df = this.getDataFormatRec(false)
            if (df != null)
                ret = df.shape.toInt()
            return ret
        }
        set(shape) {
            val df = this.getDataFormatRec(true)
            df!!.setShape(shape)
        }

    /**
     * returns true if this series has smoothed lines
     *
     * @return
     */
    val hasSmoothedLines: Boolean
        get() {
            val df = this.getDataFormatRec(false)
            return df?.smoothedLines ?: false
        }

    /**
     * retrieve the series/bar color for this series
     * NOTE: for Pie Charts, must use getPieSliceColor
     *
     * @return color int
     * @see getPieSliceColor
     */
    // otherwise, color is automatic or default chart series color
    val seriesColor: String
        get() {
            val df = this.getDataFormatRec(false)
            val type = this.parentChart!!.chartType
            val seriesNumber = df!!.seriesNumber.toInt()
            if (type == ChartConstants.PIECHART)
                return FormatHandle.colorToHexString(FormatHandle.COLORTABLE[automaticSeriesColors[seriesNumber]])
            val bg = df.bgColor
            return bg ?: FormatHandle.colorToHexString(FormatHandle.COLORTABLE[automaticSeriesColors[seriesNumber]])
        }

    /**
     * return the type of markers for each series:
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
    val markerFormat: Int
        get() {
            val df = this.getDataFormatRec(false)
            return df?.markerFormat ?: 0
        }


    /**
     * return data label options as an int
     * <br></br>can be one or more of:
     * <br></br>SHOWVALUE= 0x1;
     * <br></br>SHOWVALUEPERCENT= 0x2;
     * <br></br>SHOWCATEGORYPERCENT= 0x4;
     * <br></br>SMOOTHEDLINE= 0x8;
     * <br></br>SHOWCATEGORYLABEL= 0x10;
     * <br></br>SHOWBUBBLELABEL= 0x20;
     * <br></br>SHOWSERIESLABEL= 0x40;
     *
     * @return a combination of data label options above or 0 if none
     * @see AttachedLabel
     */
    // Extended Label -- add to attachedlabel, if any
    val dataLabel: Int
        get() {
            var datalabels = 0
            val dl = Chart.findRec(this.chartArr, DataLabExtContents::class.java) as DataLabExtContents
            if (dl != null) {
                datalabels = dl.typeInt
            }
            val df = this.getDataFormatRec(false)
            if (df != null) {
                datalabels = datalabels or df.dataLabelTypeInt
            }
            return datalabels
        }

    /**
     * return OOXML dPt (data points) for this series
     *
     * @return
     */
    val dPt: Array<DPt>?
        get() = if (dPts != null) dPts!!.toTypedArray() as Array<DPt> else null

    override fun init() {
        super.init()
        sdtX = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        sdtY = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt()).toInt()
        cValx = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt()).toInt()
        cValy = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt()).toInt()
        sdtBSz = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt()).toInt()
        sdtValBSz = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt()).toInt()
        if (DEBUGLEVEL > 10) Logger.logInfo(toString())
    }

    fun update() {
        val rkdata = this.data
        var b = ByteTools.shortToLEBytes(sdtX.toShort())
        rkdata[0] = b[0]
        rkdata[1] = b[1]
        b = ByteTools.shortToLEBytes(sdtY.toShort())
        rkdata[2] = b[0]
        rkdata[3] = b[1]
        b = ByteTools.shortToLEBytes(cValx.toShort())
        rkdata[4] = b[0]
        rkdata[5] = b[1]
        b = ByteTools.shortToLEBytes(cValy.toShort())
        rkdata[6] = b[0]
        rkdata[7] = b[1]
        b = ByteTools.shortToLEBytes(sdtBSz.toShort())
        rkdata[8] = b[0]
        rkdata[9] = b[1]
        b = ByteTools.shortToLEBytes(sdtValBSz.toShort())
        rkdata[10] = b[0]
        rkdata[11] = b[1]
        this.data = rkdata
    }

    /**
     * sets the legend for this series to a text value
     *
     * @param newLegend new text value for legend for the current series
     * @param wbh       workbookhandle
     */
    fun setLegend(newLegend: String, wbh: WorkBookHandle) {
        this.legendAi!!.setLegend(newLegend)
        val parent = this.parentChart
        parent!!.chartSeries.legends = null    // ensure cache is cleared
        parent.legend!!.adjustWidth(parent.getMetrics(wbh), parent.chartType, parent.chartSeries.getLegends())
    }

    fun hasBubbleSizes(): Boolean {
        return sdtValBSz > 0
    }

    /**
     * Gets the dataformat record associated with this Series, if any.
     * If none present, option to create a basic DataFormat set of records DataFormat controls
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
     * retrieves the data format record which corresponds to the desired pie slice
     *
     * @param slice
     * @return
     */
    private fun getDataFormatRecSlice(slice: Int, bCreate: Boolean): DataFormat? {
        var df = this.getDataFormatRec(false)
        if (df == null) {
            Logger.logErr("Series.getDataFormatRecSlice: cannot find data format record")
            return null
        }
        val seriesNumber = df.seriesNumber.toInt()
        // for PIE charts, DataFormats are stored in 1st series,
        // just after the initial data format rec
        // must check point number to see if it's the desired df
        val s: Series
        if (seriesNumber == 0)
        // we're on the first one
            s = this
        else {    // should not have more than 1 series for a pie chart!
            s = parentChart!!.allSeries[seriesNumber] as Series
        }
        var i = Chart.findRecPosition(s.chartArr, DataFormat::class.java) // get position of the first df
        i++    // skip 1st
        var lastSlice = 0
        while (i < s.chartArr.size) {
            if (s.chartArr[i] is DataFormat) {
                df = s.chartArr[i] as DataFormat
                lastSlice = df.pointNumber.toInt()
                if (df.pointNumber.toInt() == slice)
                    return df
            }
            i++
        }
        // create
        if (bCreate) {
            i--
            while (lastSlice <= slice) {
                df = DataFormat.getPrototypeWithFormatRecs(this.parentChart) as DataFormat
                df.setPointNumber(lastSlice++)
                df.parentChart = this.parentChart
                s.chartArr.add(i++, df)
            }
            return df
        }
        return null
    }

    /**
     * set smooth lines setting (applicable for line, scatter charts)
     *
     * @param smooth
     */
    fun setHasSmoothLines(smooth: Boolean) {
        val df = this.getDataFormatRec(true)
        df!!.setSmoothLines(smooth)
    }

    /**
     * sets this series to have lines
     * <br></br>Style of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
     */
    fun setHasLines(lineStyle: Int) {
        val df = this.getDataFormatRec(true)
        df!!.setHasLines(lineStyle)
    }

    /**
     * sets the color for this series
     * NOTE: for PIE Charts, use setPieSliceColor
     *
     * @param clr String color hex String
     * @see setPieSliceColor
     */
    fun setColor(clr: String) {
        val type = this.parentChart!!.chartType
        if (type == ChartConstants.PIECHART) {
            setPieSliceColor(clr, this.seriesIndex)
            return
        }
        val df = this.getDataFormatRec(true)
        df!!.setSeriesColor(clr)
    }

    /**
     * if the exact/correct color index is not used, the fill color comes out black
     *
     * @param clr
     * @return
     */
    private fun ensureCorrectColorInt(clr: Int): Int {
        var clr = clr
        // "The Chart color table is a subset of the full color table"
        if (clr == FormatConstants.COLOR_RED)
            clr = FormatConstants.COLOR_RED_CHART
        if (clr == FormatConstants.COLOR_BLUE)
            clr = FormatConstants.COLOR_BLUE_CHART
        if (clr == FormatConstants.COLOR_YELLOW)
            clr = FormatConstants.COLOR_YELLOW_CHART
        if (clr == FormatConstants.COLOR_DARK_GREEN)
        // no standard chart dark green ...
            clr = FormatConstants.COLOR_GREEN
        if (clr == FormatConstants.COLOR_DARK_YELLOW)
        // no standard chart dark yellow ...
            clr = FormatConstants.COLOR_YELLOW_CHART
        if (clr == FormatConstants.COLOR_OLIVE_GREEN)
            clr = FormatConstants.COLOR_OLIVE_GREEN_CHART
        if (clr == FormatConstants.COLOR_WHITE)
            clr = FormatConstants.COLOR_WHITE3
        return clr

    }

    /**
     * sets the color for this series
     * NOTE: for PIE Charts, use setPieSliceColor
     *
     * @param clr color int
     * @see setPieSliceColor
     */
    fun setColor(clr: Int) {
        var clr = clr
        clr = ensureCorrectColorInt(clr)
        val type = this.parentChart!!.chartType
        if (type == ChartConstants.PIECHART) {
            setPieSliceColor(clr, this.seriesIndex)
            return
        }
        val df = this.getDataFormatRec(true)
        df!!.setSeriesColor(clr)

    }


    /**
     * sets the color of the desired pie slice
     *
     * @param clr   color int
     * @param slice 0-based pie slice number
     */
    fun setPieSliceColor(clr: Int, slice: Int) {
        var clr = clr
        clr = ensureCorrectColorInt(clr)
        val type = this.parentChart!!.chartType
        if (type != ChartConstants.PIECHART)
            return
        val df = this.getDataFormatRecSlice(slice, true)
        if (df != null)
            df.setPieSliceColor(clr, slice)
        else
            Logger.logErr("Series.setPieSliceColor: unable to fnd pie slice record")
    }


    /**
     * sets the color of the desired pie slice
     *
     * @param clr   color int
     * @param slice 0-based pie slice number
     */
    fun setPieSliceColor(clr: String, slice: Int) {
        val type = this.parentChart!!.chartType
        if (type != ChartConstants.PIECHART)
            return
        val df = this.getDataFormatRecSlice(slice, true)
        if (df != null)
            df.setPieSliceColor(clr, slice)
        else
            Logger.logErr("Series.setPieSliceColor: unable to fnd pie slice record")
    }


    /**
     * get the pie slice color in this pie chart
     *
     * @param slice
     * @return color int
     */
    fun getPieSliceColor(slice: Int): String? {
        val type = this.parentChart!!.chartType
        if (type != ChartConstants.PIECHART)
            return null
        val df = this.getDataFormatRecSlice(slice, false)
        if (df != null) {
            val bg = df.bgColor
            if (bg != null)
                return bg
        }
        // otherwise, color is automatic or default chart series color
        return FormatHandle.colorToHexString(FormatHandle.COLORTABLE[automaticSeriesColors[slice]])
    }

    /**
     * PIE data label information is contained within the 1st series only
     * <br></br>TODO: not implemented yet
     * return data label options as an int
     * <br></br>can be one or more of:
     * <br></br>SHOWVALUE= 0x1;
     * <br></br>SHOWVALUEPERCENT= 0x2;
     * <br></br>SHOWCATEGORYPERCENT= 0x4;
     * <br></br>SMOOTHEDLINE= 0x8;
     * <br></br>SHOWCATEGORYLABEL= 0x10;
     * <br></br>SHOWBUBBLELABEL= 0x20;
     * <br></br>SHOWSERIESLABEL= 0x40;
     *
     * @param defaultdl int default data label setting for overall chart
     * @return int array of data labels for each pie slice
     * @see AttachedLabel
     */
    fun getDataLabelsPIE(defaultdl: Int): IntArray? {
        var datalabels = 0
        val df = this.getDataFormatRec(false)
        if (df != null) {
            datalabels = datalabels or df.dataLabelTypeInt
        }
        return null
    }

    /**
     * add a dPt element (data point) for this series
     *
     * @param Sp
     */
    fun addDpt(d: DPt) {
        if (dPts == null)
            dPts = ArrayList()
        dPts!!.add(d)
    }

    /**
     * returns the val OOXML element that defines the values for the series values
     *
     * @param valstr either "val" or "yval" for scatter/bubble charts
     * @return
     */
    fun getValOOXML(valstr: String): StringBuffer {
        val ooxml = StringBuffer()
        ooxml.append("<c:$valstr>")
        ooxml.append("\r\n")

        var seriesAi: Ai? = null
        for (i in chartArr.indices) {
            val br = chartArr[i]
            if (br.opcode == XLSConstants.AI) {
                val thisAi = br as Ai
                if (thisAi.type == Ai.TYPE_VALS) {
                    seriesAi = thisAi
                    break
                }

            }
        }

        ooxml.append("<c:numRef>")
        ooxml.append("\r\n")        // number reference
        ooxml.append("<c:f>" + OOXMLAdapter.stripNonAscii(seriesAi!!.toString()) + "</c:f>")
        ooxml.append("\r\n")    // string range
        // Need numCache for chart lines apparently
        ooxml.append("<c:numCache>")
        ooxml.append("\r\n")        // specifies the last data shown on the chart for a series
        // formatCode	== format pattern
        ooxml.append("<c:formatCode>" + this.seriesFormatPattern + "</c:formatCode>")
        val cr = CellRange(seriesAi.toString(), parentChart!!.wbh, false)
        val ch = cr.getCells()
        // ptCount	== point count
        ooxml.append(getValueRangeOOXML(ch))
        // pt * n  == a Numeric Point each has a <v> child, an idx attribute and an optional formatcode attribute
        ooxml.append("</c:numCache>")
        ooxml.append("\r\n")        //

        ooxml.append("</c:numRef>")
        ooxml.append("\r\n")
        ooxml.append("</c:$valstr>")
        ooxml.append("\r\n")
        return ooxml

    }

    /**
     * return the cat OOXML element used to define a series category
     *
     * @param cat    string category cell range for the given series --
     * almost always the same for each series (except for scatter/bubble charts)
     * @param catstr either "cat" or "xval" for scatter/bubble charts
     * cat elements must contain string references
     * xval contain numeric references
     * @return
     */
    fun getCatOOXML(cat: String, catstr: String): StringBuffer {
        val ooxml = StringBuffer()
        if ("" != cat) {  // causes 1004 vb error upon Excel SAVE - BAXTER SAVE BUG
            ooxml.append("<c:$catstr>")
            ooxml.append("\r\n")            // categories contain a string "formula" ref + string caches
            if (catstr == "cat")
                ooxml.append("<c:strRef>")    // string reference
            else
                ooxml.append("<c:numRef>")    // number reference
            ooxml.append("\r\n")
            ooxml.append("<c:f>" + OOXMLAdapter.stripNonAscii(cat) + "</c:f>")
            ooxml.append("\r\n")
            /* 20090211 KSC: if errors in referenced cells whole chart will error; best to avoid caching at all */
            if (catstr == "cat")
                ooxml.append("</c:strRef>")    // string reference
            else
                ooxml.append("</c:numRef>")    // number reference
            ooxml.append("\r\n")
            ooxml.append("</c:$catstr>")
            ooxml.append("\r\n")
        } else {
            // TESTING-- remove when done
            //Logger.logWarn("ChartHandle.getOOXML: null category found- skipping");
        }
        return ooxml
    }

    /**
     * returns the bubbleSize OOXML element that defines the values for the series values
     *
     * @param isBubble3d true if it's a 3d bubble chart
     * @return
     */
    fun getBubbleOOXML(isBubble3d: Boolean): StringBuffer {
        val ooxml = StringBuffer()
        var bubbleAi: Ai? = null
        for (i in chartArr.indices) {
            val br = chartArr[i]
            if (br.opcode == XLSConstants.AI) {
                val thisAi = br as Ai
                if (thisAi.type == Ai.TYPE_BUBBLES) {
                    bubbleAi = thisAi
                    break
                }
            }
        }

        ooxml.append("<c:bubbleSize>")
        ooxml.append("\r\n")
        ooxml.append("<c:numRef>")
        ooxml.append("\r\n")        // number reference
        ooxml.append("<c:f>" + bubbleAi!!.toString() + "</c:f>")
        ooxml.append("\r\n")
        ooxml.append("<c:numCache>")
        ooxml.append("\r\n")
        try {
            val cells = CellRange.getCells(bubbleAi.toString(), parentChart!!.wbh)
            ooxml.append(getValueRangeOOXML(cells))
            ooxml.append("\r\n")
        } catch (e: NumberFormatException) {
            Logger.logErr("geteriesOOXML: Number format exception for Bubble Range: $bubbleAi")
        }

        ooxml.append("</c:numCache>")
        ooxml.append("\r\n")
        ooxml.append("</c:numRef>")
        ooxml.append("\r\n")
        ooxml.append("</c:bubbleSize>")
        ooxml.append("\r\n")
        if (isBubble3d)
            ooxml.append("<c:bubble3D val=\"1\"/>")
        ooxml.append("\r\n")
        return ooxml
    }

    /**
     * returns the bubbleSize OOXML element that defines the values for the series values
     *
     * @param isBubble3d true if it's a 3d bubble chart
     * @return
     */
    fun getLegendOOXML(from2003: Boolean): StringBuffer {
        val ooxml = StringBuffer()
        val txt = this.legendText
        val ai = this.legendAi
        /*       String txt= null;
       try {
       	io.starter.formats.XLS.formulas.Ptg[] p= ai.getCellRangePtgs();
       	txt= ((io.starter.formats.XLS.formulas.PtgRef)p[0]).getFormattedValue();
       } catch (Exception e) {}
       try {
           if (chartArr.size()>i+1) {
              SeriesText st = (SeriesText)chartArr.get(i+1);
              if (st!=null)
           	   txt= st.toString();
           }
       }catch(ClassCastException e) {
           // couldn't find it!
       }
*/

        ooxml.append("<c:tx>")
        ooxml.append("\r\n")
        ooxml.append("<c:strRef>")
        ooxml.append("\r\n")        // string reference
        if (ai != null) {
            ooxml.append("<c:f>" + OOXMLAdapter.stripNonAscii(ai.definition) + "</c:f>")
            ooxml.append("\r\n")
            ooxml.append("<c:strCache>")
            ooxml.append("\r\n")
            ooxml.append("<c:ptCount val=\"1\"/>")
            ooxml.append("\r\n")
            ooxml.append("<c:pt idx=\"0\">")
            ooxml.append("\r\n")
            ooxml.append("<c:v>" + OOXMLAdapter.stripNonAscii(txt) + "</c:v>")
            ooxml.append("</c:pt>")
            ooxml.append("\r\n")
            ooxml.append("</c:strCache>")
            ooxml.append("\r\n")
            ooxml.append("</c:strRef>")
            ooxml.append("\r\n")
            ooxml.append("</c:tx>")
            ooxml.append("\r\n")
            if (this.spPr != null)
                ooxml.append(this.spPr!!.ooxml)
            else if (from2003) {
                val ss: SpPr
                if (parentChart!!.chartType != ChartConstants.RADARCHART)
                    ss = SpPr("c", this.seriesColor.substring(1), 12700, "000000")
                else
                    ss = SpPr("c", null, 25400, this.seriesColor.substring(1))
                ooxml.append(ss.ooxml)
            }
        }
        return ooxml
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 7290108485347063887L
        var SERIES_TYPE_NUMERIC = 1
        var SERIES_TYPE_STRING = 3

        /**
         * Create a series with all sub components.
         *
         *
         * REcords are...
         * AI
         * SeriesText (optional?)
         * AI
         * AI
         * AI
         * DataFormat
         * SerToCrt
         *
         * @param seriesData
         * @return
         */
        fun getPrototype(seriesRange: String, categoryRange: String, bubbleRange: String, legendRange: String?, legendText: String, chartobj: ChartType): Series {
            val series = Series.prototype as Series?
            val parentChart = chartobj.parentChart
            val book = parentChart!!.workBook
            var ai: Ai
            // create Series text with Legend
            if (legendRange != null) {
                ai = Ai.getPrototype(Ai.AI_TYPE_LEGEND) as Ai
                ai.workBook = book
                ai.setSheet(parentChart.sheet)    // 20080124 KSC: changeAiLocation from "A1" to ""
                try {
                    ai.changeAiLocation("", legendRange/*(seriesText.getWorkSheetName() + "!" + seriesText.getCellAddress())*/)
                } catch (e: Exception) {
                }
                // it's OK to not have a valid range
                series!!.addChartRecord(ai)
                val st = SeriesText.getPrototype(legendText)
                ai.setSeriesText(st)
                series.addChartRecord(st)
                // 20091102 KSC: when adding series legend will not expand correctly if autopositioning is turned off
                // [BugTracker 2844]
                val l = chartobj.dataLegend
                if (l != null) {
                    l.setAutoPosition(true)
                    l.incrementHeight(parentChart.coords[3].toFloat())
                }

            } else {
                ai = Ai.getPrototype(Ai.AI_TYPE_NULL_LEGEND) as Ai
                ai.workBook = book
                ai.setSheet(parentChart.sheet)
                series!!.addChartRecord(ai)
            }
            //        parentChart.addAi(ai);
            // create Series Value Ai
            ai = Ai.getPrototype(Ai.AI_TYPE_SERIES) as Ai
            ai.parentChart = parentChart
            ai.workBook = book
            ai.setSheet(parentChart.sheet)
            try {
                ai.changeAiLocation(ai.toString(), seriesRange)
            } catch (e: Exception) {
                // not necessary to report        	Logger.logErr("Error setting Series Range: '"  + seriesRange + "'-" + e.toString());
            }
            // it's OK to not have a valid range
            series.addChartRecord(ai)
            //        parentChart.addAi(ai);
            // create Category Ai
            ai = Ai.getPrototype(Ai.AI_TYPE_CATEGORY) as Ai
            ai.workBook = book
            ai.setSheet(parentChart.sheet)
            try {
                ai.changeAiLocation(ai.toString(), categoryRange)
            } catch (e: Exception) {
            }
            // it's OK to not have a valid range
            series.addChartRecord(ai)
            //        parentChart.addAi(ai);
            // create Bubble (undocumented) Ai
            ai = Ai.getPrototype(Ai.AI_TYPE_BUBBLE) as Ai
            ai.workBook = book
            ai.setSheet(parentChart.sheet)
            if (bubbleRange != "") {
                try {
                    ai.changeAiLocation(ai.toString(), bubbleRange)
                } catch (e: Exception) {
                }
                // it's OK to not have a valid range
                ai.setRt(2)
                /* 20120123 KSC: confused vs. per-series format verses format in chartformat rec ...
 *              if (((BubbleChart) parentChart.getChartObject()).is3d()) {

            }*/
            }
            series.addChartRecord(ai)
            var df: DataFormat? = null
            df = DataFormat.prototype as DataFormat?
            // update the data format correctly
            val ser = parentChart.allSeries    // get ALL series
            var yi = -1        // Changed from 0
            var iss = -1    // ""
            for (i in ser.indices) {
                val srs = ser[i] as Series
                val newYi = srs.seriesIndex
                val newIss = srs.seriesNumber
                if (newYi > yi) yi = newYi
                if (newIss > iss) iss = newIss
            }
            yi++
            iss++
            df!!.setSeriesIndex(yi)
            df.setSeriesNumber(iss)
            if (chartobj.barShape != 0) { // must ensure each series contains proper shape records
                df.setShape(chartobj.barShape)
            }
            series.addChartRecord(df)
            val stc = SerToCrt.prototype as SerToCrt?
            // get the correct chart index for the sertocrt
            var vCount = 0
            var cCount = 0
            var bCount = 0
            if (ser.size > 0) {    // 20070709 KSC: will be 0 if adding new blank chart
                val s = ser[0] as Series
                val cr = s.chartRecords
                for (i in cr.indices) {
                    val b = cr[i] as BiffRec
                    if (b.opcode == XLSConstants.SERTOCRT) {
                        val stcc = b as SerToCrt
                        stc!!.data = stcc.data
                    }
                }
                // set the series level variables correctly
                vCount = s.valueCount
                cCount = s.categoryCount
                bCount = s.bubbleCount
            }
            series.init()
            // 20070711 KSC: vCount and cCount are via current range, no??
            try {
                if (seriesRange.indexOf(":") != -1) {
                    val coords = io.starter.OpenXLS.ExcelTools.getRangeCoords(seriesRange)
                    vCount = coords[4]
                } else {
                    vCount = 1
                }
                series.valueCount = vCount
            } catch (e: Exception) {
            }

            try {
                cCount = io.starter.OpenXLS.ExcelTools.getRangeCoords(categoryRange)[4]
                series.categoryCount = cCount
            } catch (e: Exception) {
            }

            if (bubbleRange != "") {
                try {
                    bCount = io.starter.OpenXLS.ExcelTools.getRangeCoords(bubbleRange)[4]
                    series.bubbleCount = bCount
                } catch (e: Exception) {
                }

            }
            series.addChartRecord(stc)
            return series
        }

        val prototype: XLSRecord?
            get() {
                val s = Series()
                s.opcode = XLSConstants.SERIES
                s.data = s.PROTOTYPE_BYTES
                s.init()
                return s
            }

        // Periwinkle 	Plum+ 	Ivory 	Light Turquoise 	Dark Purple 	Coral 	Ocean Blue 	Ice Blue  {17, 25, 19, 27, 28, 22, 23, 24};
        // try these color int numbers instead:
        // alternative explanation:  chart fills 16-23, chart lines 24-31
        var automaticSeriesColors = intArrayOf(24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62)        // also used in mapping default colors in AreaFormat, MarkerFormat, Frame ...

        /**
         * generate the OOXML used to represent this set of value cells (element numRef)
         *
         * @param cells
         * @return
         */
        private fun getValueRangeOOXML(cells: Array<CellHandle>?): String {
            val ooxml = StringBuffer()
            if (cells == null) return "<c:ptCount val=\"0\"/>"
            ooxml.append("<c:ptCount val=\"" + cells.size + "\"/>")
            ooxml.append("\r\n")
            for (j in cells.indices) {
                ooxml.append("<c:pt idx=\"$j\">")
                ooxml.append("\r\n")
                if (cells[j].stringVal != "NaN") {
                    ooxml.append("<c:v>" + cells[j].stringVal + "</c:v>")
                    ooxml.append("\r\n")
                } else {    // appears that NaN is an invalid entry
                    ooxml.append("<c:v>0</c:v>")
                    ooxml.append("\r\n")
                }
                ooxml.append("</c:pt>")
                ooxml.append("\r\n")
            }
            return ooxml.toString()
        }
    }


}