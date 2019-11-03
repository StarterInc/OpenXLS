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
import io.starter.OpenXLS.ExcelTools
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.OOXML.DLbls
import io.starter.formats.OOXML.DPt
import io.starter.formats.OOXML.Marker
import io.starter.formats.OOXML.SpPr
import io.starter.formats.XLS.*
import io.starter.formats.XLS.formulas.Ptg
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger
import org.json.JSONArray
import org.json.JSONException
import org.json.JSONObject
import org.xmlpull.v1.XmlPullParser

import java.io.Serializable
import java.util.ArrayList
import java.util.HashMap
import java.util.Stack
import java.util.Vector

class ChartSeries : ChartConstants, Serializable {
    private val series = ArrayList()    // ALL series in chart MAPPED to chart that "owns" it
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
    var seriesJSON: JSONArray? = null
        @Throws(JSONException::class)
        set(s) {
            field = JSONArray(s.toString())
        }            // save series JSON for update comparisons later
    protected var minmaxcache: DoubleArray? = null    // stores minimum/maximum/minor/major for chart scale; cached;
    var legends: Array<String>? = null
    protected var seriesranges: ArrayList<*>? = null
    protected var categories: Array<Any>? = null
    protected var seriesvalues: ArrayList<*>? = null
    protected var seriescolors: Array<String>? = null
    protected var parentChart: Chart

    /**
     * get all the series objects for ALL charts
     * (i.e. even in overlay charts)
     *
     * @return
     */
    val allSeries: Vector<*>
        get() = getAllSeries(-1)

    /**
     * retrieve the current data range JSON for comparisons
     *
     * @return JSONObject in form of: c:catgeory range, {v:series range, l:legendrange, b: bubble sizes} for each series
     */
    //v=series val range, l=legend cell [,b= bubble sizes]
    // keep going
    // seriesJSON.getJSONArray("Series").length()
    // seriesJSON.getJSONArray("Series").get(1)
    // ((JSONObject)seriesJSON.getJSONArray("Series").get(1)).get("v")	value range
    // ((JSONObject)seriesJSON.getJSONArray("Series").get(1)).get("l")	legend ref
    // ((JSONObject)seriesJSON.getJSONArray("Series").get(1)).has("b")	bubble sizes
    // seriesJSON.get("c")		category range
    val dataRangeJSON: org.json.JSONObject
        get() {
            val seriesJSON = JSONObject()
            val allseries = this.getAllSeries(-1)
            try {
                val series = JSONArray()
                for (i in allseries.indices) {
                    try {
                        val thisseries = allseries[i] as Series
                        val serAi = thisseries.seriesValueAi
                        if (i == 0) {
                            val catAi = thisseries.categoryValueAi
                            seriesJSON.put("c", catAi!!.toString())
                        }
                        val seriesvals = JSONObject()
                        seriesvals.put("v", serAi!!.toString())
                        seriesvals.put("l", thisseries.legendRef)
                        if (thisseries.hasBubbleSizes())
                            seriesvals.put("b", thisseries.bubbleValueAi!!.toString())
                        series.put(seriesvals)
                    } catch (e: Exception) {
                    }

                }
                seriesJSON.put("Series", series)
            } catch (e: JSONException) {
                Logger.logErr("ChartSeries.getDataRangeJSON:  Error retrieving Series Information: $e")
            }

            return seriesJSON
        }

    /**
     * return an array of ALL cell references of the chart
     */
    val cellRangePtgs: Array<Ptg>
        get() {
            val locptgs = CompatibleVector()
            for (i in series.indices) {
                val s = series.get(i)[0] as Series
                for (j in s.chartArr.indices) {
                    val br = s.chartArr[j]
                    if (br.opcode == XLSConstants.AI) {
                        try {
                            val ps = (br as Ai).cellRangePtgs
                            for (t in ps.indices) locptgs.add(ps[t])
                        } catch (e: Exception) {
                        }

                    }
                }
            }
            val ret = arrayOfNulls<Ptg>(locptgs.size)
            locptgs.toTypedArray()
            return ret
        }

    /**
     * @return an HashMap of Series Range Ptgs mapped by Series representing all the series in the chart
     */
    // we're done, move onto next
    val seriesPtgs: HashMap<*, *>
        get() {
            val seriesPtgs = HashMap()
            for (i in series.indices) {
                val s = series.get(i)[0] as Series
                for (j in s.chartArr.indices) {
                    val br = s.chartArr[j]
                    if (br.opcode == XLSConstants.AI) {
                        if ((br as Ai).type == Ai.TYPE_VALS) {
                            try {
                                val ps = br.cellRangePtgs
                                seriesPtgs.put(s, ps)
                                break
                            } catch (e: Exception) {
                            }

                        }
                    }
                }
            }
            return seriesPtgs
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
    val markerFormats: IntArray
        get() {
            val markers = IntArray(series.size)
            for (i in series.indices) {
                val s = series.get(i)[0] as Series
                markers[i] = s.markerFormat
            }
            return markers
        }

    internal var seriesNumber = 0

    val seriesRanges: ArrayList<*>?
        get() {
            if (seriesranges == null)
                this.getMetrics(true)
            return seriesranges
        }

    val seriesValues: ArrayList<*>?
        get() {
            if (seriesvalues == null)
                this.getMetrics(true)
            return seriesvalues
        }

    val seriesBarColors: Array<String>?
        get() {
            if (seriescolors == null)
                this.getMetrics(true)
            return seriescolors
        }
    //	protected transient WorkBookHandle wbh;

    /**
     * series stores new Object[] {Series Record, Integer.valueOf(nCharts)}
     *
     * @param o
     */
    fun add(o: Array<Any>) {
        series.add(o)
    }

    fun setParentChart(c: Chart) {
        parentChart = c
    }
    //	public void setWorkBook(WorkBookHandle wbh) { this.wbh= wbh; }

    /**
     * returns an array containing the and maximum values of Y (Value) axis, along with the
     * maximum number of series values of each series (bar, line, etc.) ...
     * <br></br>sets all series cache values:  seriesvalues, seriesranges, seriescolors, legends and minmaxcache
     *
     * <br></br>Note: this will only reset cached values if isDirty
     *
     * @return double[]
     */
    fun getMetrics(isDirty: Boolean): DoubleArray {
        if (!isDirty && minmaxcache != null) return minmaxcache
        // trap minimum, maximum + number of series
        val co = parentChart.chartObject    // default chart object TODO: overlay charts
        seriesvalues = ArrayList()
        seriesranges = ArrayList()
        val sht = parentChart.sheet
        val s = this.getAllSeries(-1)
        // Category values *******************************************************************************
        if (s.size > 0) {
            try {
                var cr = CellRange((s[0] as Series).categoryValueAi!!.toString(), parentChart.wbh, true)
                val ch = cr.getCells()
                if (ch != null) { // found a template with a chart with series but no categories
                    categories = arrayOfNulls(ch.size)
                    for (j in ch.indices) {
                        try {
                            categories[j] = ch[j].getFormattedStringVal(true)
                        } catch (e: IllegalArgumentException) { // catch format exceptions
                            categories[j] = ch[j].stringVal
                        }

                    }
                } else if (s.size > 0) {
                    cr = CellRange((s[0] as Series).seriesValueAi!!.toString(), parentChart.wbh, true)
                    val sz = cr.getCells()!!.size
                    categories = arrayOfNulls(sz)
                    for (j in 0 until sz)
                        categories[j] = Integer.valueOf(j + 1)!!.toString()
                }
            } catch (e: Exception) {
                Logger.logWarn("ChartSeries.getMinMax: $e")
            }

        }
        // Series colors, labels and values ***************************************************************
        var yMax = 0.0
        var yMin = java.lang.Double.MAX_VALUE
        var nseries = 0
        seriescolors = null
        legends = null
        val charttype = co.chartType
        // obtain/store series colors, store series values and trap maximum and
        // minimun values so can be used below for axis scale
        /*
         * A Scatter chart has two value axes, showing one set of numerical data along the x-axis and another along the y-axis.
         * It combines these values into single data points and displays them in uneven intervals, or clusters
         */
        if (charttype != ChartConstants.PIECHART && charttype != ChartConstants.DOUGHNUTCHART) {
            seriescolors = arrayOfNulls(s.size)
            legends = arrayOfNulls(s.size)
            for (i in s.indices) {
                val myseries = s[i] as Series
                seriescolors[i] = myseries.seriesColor
                legends[i] = io.starter.formats.XLS.OOXMLAdapter.stripNonAscii(myseries.legendText).toString()
                val cr = CellRange(myseries.seriesValueAi!!.toString(), parentChart.wbh, true)
                val ch = cr.getCells()
                nseries = Math.max(nseries, ch!!.size)
                val seriesvals: DoubleArray
                val sranges: Array<String>
                //				String[] series_strings;
                if (!myseries.hasBubbleSizes()) {
                    seriesvals = DoubleArray(nseries)
                    sranges = arrayOfNulls(nseries)
                } else {
                    seriesvals = DoubleArray(nseries * 2)
                    sranges = arrayOfNulls(nseries * 2)
                }
                //				series_strings = new String[seriesvals.length];

                for (j in ch.indices) {
                    try {
                        sranges[j] = ch[j].cellAddressWithSheet
                        seriesvals[j] = ch[j].doubleVal
                        if (java.lang.Double.isNaN(seriesvals[j]))
                            seriesvals[j] = 0.0
                        yMax = Math.max(yMax, seriesvals[j])
                        yMin = Math.min(yMin, seriesvals[j])
                    } catch (n: NumberFormatException) {
                    }

                }
                if (myseries.hasBubbleSizes()) { // append bubble sizes to series values ... see BubbleChart.getSVG for parsing
                    val z = ch.size
                    val crb = CellRange(myseries.bubbleValueAi!!.toString(), parentChart.wbh, true)
                    val chb = crb.getCells()
                    for (j in ch.indices) {
                        seriesvals[j + z] = chb!![j].doubleVal
                        sranges[j + z] = chb[j].cellAddressWithSheet
                    }
                }
                seriesvalues!!.add(seriesvals)   // trap and add series value points
                seriesranges!!.add(sranges)        // trap series range
            }
        } else if (charttype == ChartConstants.DOUGHNUTCHART && s.size > 1) { // like a PIE chart but can have multiple series
            legends = arrayOfNulls(categories!!.size)        // for PIE/DONUT charts, legends are actually category labels, not series labels
            for (i in categories!!.indices)
                legends[i] = io.starter.formats.XLS.OOXMLAdapter.stripNonAscii(categories!![i].toString()).toString()
            for (i in s.indices) {
                val myseries = s[i] as Series
                // legends[i]=
                // io.starter.formats.XLS.OOXMLAdapter.stripNonAscii(myseries.getLegend());
                val cr = CellRange(myseries.seriesValueAi!!.toString(), parentChart.wbh, true)
                val ch = cr.getCells()
                val seriesvals = DoubleArray(ch!!.size)
                val sranges = arrayOfNulls<String>(ch.size)
                if (seriescolors == null)
                    seriescolors = arrayOfNulls(ch.size)
                for (j in ch.indices) {
                    try {
                        seriesvals[j] = ch[j].doubleVal
                        if (ch[j].workSheetHandle!!.mysheet == sht)
                            sranges[j] = ch[j].cellAddress
                        yMax = Math.max(yMax, seriesvals[j])
                        yMin = Math.min(yMin, seriesvals[j])
                        if (i == 0) { // only do for 1st series; will be the
                            // same for rest
                            seriescolors[j] = myseries.getPieSliceColor(j)
                            /*if (seriescolors[j] == 0x4D
									|| seriescolors[j] == 0x4E)
								seriescolors[j] = io.starter.formats.XLS.FormatConstants.COLOR_WHITE;*/
                        }

                    } catch (n: NumberFormatException) {
                    }

                }
                seriesvalues!!.add(seriesvals) // trap and add series value points
                seriesranges!!.add(sranges)        // trap series range
            }
        } else { // PIES - only 1 series
            if (s.size > 0) {
                // PIE: 1 series data
                val cats = CellRange((s[0] as Series).categoryValueAi!!.toString(), parentChart.wbh, true).getCells()
                if (cats != null) {
                    nseries = cats.size
                    legends = arrayOfNulls(cats.size) // for PIE charts, legends are actually category labels, not series labels
                    for (i in cats.indices)
                        legends[i] = cats[i].getFormattedStringVal(true)
                }
                seriescolors = arrayOfNulls(nseries)
                val myseries = s[0] as Series
                try {
                    val cr = CellRange(myseries.seriesValueAi!!.toString(), parentChart.wbh, true)
                    val ch = cr.getCells()
                    // error trap - shouldn't happen
                    if (ch!!.size != nseries) {
                        Logger.logWarn("ChartHandle.getSeriesInfo: unexpected Pie Chart structure")
                        nseries = Math.min(nseries, ch.size)
                    }
                    val seriesvals = DoubleArray(nseries)
                    val sranges = arrayOfNulls<String>(nseries)
                    for (i in 0 until nseries) {
                        seriescolors[i] = myseries.getPieSliceColor(i)
                        /*if (seriescolors[i] == 0x4D || seriescolors[i] == 0x4E)
							seriescolors[i] = io.starter.formats.XLS.FormatConstants.COLOR_WHITE;*/
                        // legends[i]=
                        // io.starter.formats.XLS.OOXMLAdapter.stripNonAscii(myseries.getLegend());
                        // // same for every series ...
                        seriesvals[i] = ch[i].doubleVal
                        if (ch[i].workSheetHandle!!.mysheet == sht)
                            sranges[i] = ch[i].cellAddress
                        yMax = Math.max(yMax, seriesvals[i])
                        yMin = Math.min(yMin, seriesvals[i])
                    }
                    seriesvalues!!.add(seriesvals) // trap and add series value points
                    seriesranges!!.add(sranges)        // trap series range
                } catch (e: IllegalArgumentException) {
                    // error in cell range sheet ...
                }

            }
        }
        // For stacked-type charts, must sum values for ymax
        if (co.isStacked) {
            // scale is SUM of values, yMax is maximum total per series point
            val sum = DoubleArray(nseries)
            for (i in seriesvalues!!.indices) {
                val seriesv = seriesvalues!![i] as DoubleArray
                for (j in seriesv.indices) {
                    sum[j] = sum[j] + seriesv[j]
                }
            }
            yMax = 0.0
            for (i in 0 until nseries) {
                yMax = Math.max(sum[i], yMax)
            }
        }
        minmaxcache = DoubleArray(2)
        minmaxcache[0] = yMin
        minmaxcache[1] = yMax
        //    	minmaxcache[2]= nSeries;
        //    	minmaxcache= new double[3][];	// 3 possible axes: x, y and z
        //    	minmaxcache[axisType]= minMax;
        return minmaxcache
    }

    /**
     * Change series ranges for ALL matching series
     *
     * @param originalrange
     * @param newrange
     * @return
     */
    fun changeSeriesRange(originalrange: String, newrange: String): Boolean {
        var changed = false
        for (i in series.indices) {
            val s = series.get(i)[0] as Series
            val ai = s.seriesValueAi
            if (ai != null) {
                if (ai.toString().equals(originalrange, ignoreCase = true)) {
                    changed = ai.changeAiLocation(originalrange, newrange)
                }
            }
        }

        return changed
    }

    /**
     * Return a string representing all series in this chart
     *
     * @param nChart 0=default, 1-9= overlay charts -1 for ALL charts
     * @return
     */
    fun getSeries(nChart: Int): Array<String> {
        val seriesperchart = this.getAllSeries(nChart)
        val retStr = arrayOfNulls<String>(seriesperchart.size)
        for (i in seriesperchart.indices) {
            val s = seriesperchart[i] as Series
            val a = s.seriesValueAi
            retStr[i] = a!!.toString()
        }
        return retStr
    }

    /**
     * Return an array of strings, one for each category
     *
     * @param nChart 0=default, 1-9= overlay charts -1 for All
     * @return
     */
    fun getCategories(nChart: Int): Array<String> {
        val seriesperchart = this.getAllSeries(nChart)
        val retStr = arrayOfNulls<String>(seriesperchart.size)
        for (i in seriesperchart.indices) {
            val s = seriesperchart[i] as Series
            val a = s.categoryValueAi
            retStr[i] = a!!.toString()
        }
        return retStr
    }

    /**
     * get all the series objects in the specified chart (-1 for ALL)
     *
     * @param nChart 0=default, 1-9= overlay charts -1 for ALL series
     * @return
     */
    fun getAllSeries(nChart: Int): Vector<*> {
        /*    	if (nChart==-1) { //return all series
    		return new Vector(series.keySet()); // unordered!!!
    	}
*/
        val retVec = Vector()
        for (i in series.indices) {
            val chart = series.get(i)[1] as Int
            if (nChart == -1 || nChart == chart) {
                val s = series.get(i)[0] as Series
                retVec.add(s)
            }
        }
        return retVec
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
    fun addSeries(seriesRange: String, categoryRange: String, bubbleRange: String, legendRange: String, legendText: String, chartObject: ChartType, nChart: Int): Series {
        val s = Series.getPrototype(seriesRange, categoryRange, bubbleRange, legendRange, legendText, chartObject)
        s.parentChart = chartObject.parentChart
        s.shape = chartObject.barShape
        series.add(arrayOf<Any>(s, Integer.valueOf(nChart)))
        // Update parent chartArr
        val chartArr = s.parentChart!!.chartArr
        for (i in chartArr.indices) {
            val br = chartArr[i]
            var br2: BiffRec? = null
            if (i < chartArr.size - 1) br2 = chartArr[i + 1]
            if (br != null && (br.opcode == XLSConstants.SERIES && br2!!.opcode != XLSConstants.SERIES || br.opcode == XLSConstants.FRAME && br2!!.opcode == XLSConstants.SHTPROPS)) {
                chartArr.add(i + 1, s)
                break
            }
        }
        if (chartObject.chartType == ChartConstants.STOCKCHART)
            s.setHasLines(5)
        return s
    }

    /**
     * specialty method to take absolute index of series and remove it
     * <br></br>only used in WorkSheetHandle
     *
     * @param index
     */
    fun removeSeries(index: Int) {
        val v = this.allSeries
        this.removeSeries(v[index] as Series)
    }

    /**
     * remove desired series from chart
     *
     * @param index
     */
    fun removeSeries(seriestodelete: Series) {
        for (z in series.indices) {
            val ss = series.get(z)[0] as Series
            if (ss == seriestodelete) {
                series.removeAt(z)
                break
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
    fun getSeries(seriesName: String, nChart: Int): Series? {
        val seriesperchart = this.getAllSeries(nChart)
        for (i in seriesperchart.indices) {
            val s = seriesperchart[i] as Series
            if (s.legendText.equals(seriesName, ignoreCase = true)) return s
        }
        return null
    }

    /**
     * changes the category range which matches originalrange to new range
     *
     * @param originalrange
     * @param newrange
     * @return
     */
    fun changeCategoryRange(originalrange: String, newrange: String): Boolean {
        var changed = false
        for (i in series.indices) {
            val s = series.get(i)[0] as Series
            val ai = s.categoryValueAi
            if (ai!!.toString().equals(originalrange, ignoreCase = true)) {
                changed = ai.changeAiLocation(originalrange, newrange)
            }
        }
        return changed
    }

    /**
     * attempts to replace all category elements containing originalval text to newval
     *
     * @param originalval
     * @param newval
     * @return
     */
    fun changeTextValue(originalval: String, newval: String): Boolean {
        val changed = false
        for (i in series.indices) {
            val s = series.get(i)[0] as Series
            val ai = s.categoryValueAi
            if (ai!!.text == originalval) {
                ai.text = newval
            }

        }
        return changed
    }


    /**
     * for overlay charts, store series list
     *
     * @param nCharts
     * @param seriesList
     */
    fun addSeriesMapping(nCharts: Int, seriesList: IntArray) {
        for (i in seriesList.indices) {
            try {
                val idx = seriesList[i] - 1
                // ALL series in chart MAPPED to chart that "owns" it
                val s = series.get(idx)[0] as Series
                series.add(idx, arrayOf<Any>(s, Integer.valueOf(nCharts)))
            } catch (ae: ArrayIndexOutOfBoundsException) {
                // happens -- are they deleted series?????
            }

        }
    }

    /**
     * return an array of legend text
     *
     * @param nChart 0=default, 1-9= overlay charts -1 for ALL
     */
    fun getLegends(nChart: Int): Array<String> {
        val seriesperchart = this.getAllSeries(nChart)
        val ret = arrayOfNulls<String>(seriesperchart.size)
        for (i in seriesperchart.indices) {
            ret[i] = (seriesperchart[i] as Series).legendText
        }
        return ret
    }

    /**
     * Generate the OOXML used to represent all series for this chart type
     * @param ct        chart type
     * @return
     */
    /**
     * all contain:
     * idx			(index)
     * order		(order)
     * tx			(series text)
     * spPr		(shape properties)
     * then after, may contain:
     * bubble:  invertIfNegative, dPt, dLbls, trendline, errBars, xVal, yVal, bubbleSize, bubble3D		(bubbleChart)
     * line 	marker, dPt, dLbls, trendline, errBars, cat, val, smooth								(line3DChart, lineChart, stockChart)
     * pie:		explosion, dPt,dLbls, cat, val															(doughnutChart, ofPieChart, pie3DChart, pieChart)
     * surface: cat, val																				(surfaceChart, surface3DChart)
     * scatter: marker, dPt, dLbls, trendline, errBars, xVal, yVal, smooth								(scatterChart)
     * radar:	marker, dPt, dLbls, cat, val															(radarChart)
     * area:	pictureOptions, dPt, dLbls, trendline, errBars, cat, val								(area3DChart, areaChart)
     * bar:		invertIfNegative, pictureOptions, dPt, dLbls, trendline, errBars, cat, val, shape		(bar3DChart, barChart)
     */
    // TODO: finish options
    // TODO: refactor !!!
    fun resetSeriesNumber() {
        seriesNumber = 0
    }

    fun getOOXML(ct: Int, isBubble3d: Boolean, nChart: Int): String {
        val catstr = if (ct == ChartConstants.SCATTERCHART || ct == ChartConstants.BUBBLECHART) "xVal" else "cat"
        val valstr = if (ct == ChartConstants.SCATTERCHART || ct == ChartConstants.BUBBLECHART) "yVal" else "val"
        val ooxml = StringBuffer()

        val v = parentChart.getAllSeries(nChart)
        val defaultDL = parentChart.dataLabel
        val from2003 = !parentChart.workBook!!.isExcel2007
        val cats = this.getCategories(nChart)    // do 1x


        for (i in v.indices) {
            val s = v[i] as Series
            ooxml.append("<c:ser>")
            ooxml.append("\r\n")
            ooxml.append("<c:idx val=\"$seriesNumber\"/>")
            ooxml.append("\r\n")
            ooxml.append("<c:order val=\"" + seriesNumber++ + "\"/>")
            ooxml.append("\r\n")
            // Series Legend
            ooxml.append(s.getLegendOOXML(from2003))
            // Options for current series
            if (ct == ChartConstants.PIECHART /*&& i==0*/)
                ooxml.append("<c:explosion val=\"" + parentChart.getChartOption("Percentage") + "\"/>")
            if (s.marker != null) ooxml.append(s.marker!!.ooxml)        // only for Radar, Line or Scatter
            if (s.dPt != null) {
                val datapoints = s.dPt
                for (z in datapoints!!.indices) {
                    ooxml.append(datapoints[z].ooxml)
                }
            }
            if (s.dLbls != null)
                ooxml.append(s.dLbls!!.ooxml)
            else if (from2003) {
                val dl = s.dataLabel or defaultDL
                if (dl > 0) {    // todo: showLegendKey catpercent ? sppr + txpr
                    // TODO: spPr, txPr
                    val dlbl = DLbls(dl or 0x1 == 0x1, dl or 0x08 == 0x08, false, dl or 0x10 == 0x10, dl or 0x40 == 0x40,
                            dl or 0x2 == 0x2, dl or 0x20 == 0x20, null, null/*txpr*/)/*sppr*/
                    ooxml.append(dlbl.ooxml)
                }
            }
            if (s.hasSmoothedLines) {
                ooxml.append("<c:smooth val=\"1\"/>")
                ooxml.append("\r\n")
            }

            // Categories 							NOTE:  Categories==xVals for Scatter charts, cat for all others
            ooxml.append(s.getCatOOXML(cats[i], catstr))
            // Series ("vals")						NOTE:  Series==yVals for Scatter charts, val for all others
            ooxml.append(s.getValOOXML(valstr))    // gets the numeric data reference to define the series (values)

            if (ct == ChartConstants.BUBBLECHART) { // also include bubble sizes
                ooxml.append(s.getBubbleOOXML(isBubble3d))
            }
            ooxml.append("</c:ser>")
            ooxml.append("\r\n")
        }
        return ooxml.toString()
    }


    /**
     * if has multiple or overlay charts, update series mappings
     *
     * @param sl
     * @param thischartnumber
     */
    fun updateSeriesMappings(sl: SeriesList?, thischartnumber: Int) {
        if (sl == null)
            return     // NO seriesList record means no mapping
        val seriesmappings = ArrayList()
        for (z in series.indices) {
            val chartnumber = (series.get(z)[1] as Int).toInt()
            if (chartnumber == thischartnumber)
            // mappped to this chart
                seriesmappings.add(Integer.valueOf(z + 1))
        }
        val mappings = IntArray(seriesmappings.size)
        for (z in seriesmappings.indices) {
            mappings[z] = (seriesmappings.get(z) as Int).toInt()
        }
        try {
            sl.seriesMappings = mappings
        } catch (e: Exception) {
            throw WorkBookException("ChartSeries.updateSeriesMappings failed:$e", WorkBookException.RUNTIME_ERROR)
        }

    }

    /**
     * return Data Labels Per Series or default Data Labels, if no overrides specified
     *
     * @param defaultDL Default Data labels
     * @param charttype Chart Type Int
     * @return
     */
    fun getDataLabelsPerSeries(defaultDL: Int, charttype: Int): IntArray {
        if (charttype == ChartConstants.PIECHART || charttype == ChartConstants.DOUGHNUTCHART) {    // handled differently
            if (series.size > 0) {
                val s = series.get(0)[0] as Series
                var dls = s.getDataLabelsPIE(defaultDL)
                if (dls == null) {
                    dls = intArrayOf(defaultDL)
                }
                return dls
            }
        }
        val datalabels = IntArray(series.size)
        for (i in series.indices) {
            val s = series.get(i)[0] as Series
            datalabels[i] = s.dataLabel
            datalabels[i] = datalabels[i] or defaultDL // if no per-series setting use overall chart setting
        }
        return datalabels
    }


    // TODO: FINISH -- include cell range ... + overlay charts ...?
    fun getLegends(): Array<String>? {
        if (legends == null) {
            this.getMetrics(true)
        }
        return legends
    }


    fun getCategories(): Array<Any>? {
        if (categories == null)
            this.getMetrics(true)
        return categories
    }

    companion object {
        /**
         * serialVersionUID
         */
        private const val serialVersionUID = -7862828186455339066L


        /**
         * parse a chartSpace->chartType->ser element into our Series record/structure
         *
         * @param xpp         XML pullparser positioned at ser element
         * @param wbh         WorkBookHandle
         * @param parentChart parent chart object
         * @param lastTag
         * @return
         */
        fun parseOOXML(xpp: XmlPullParser, wbh: WorkBookHandle, parentChart: ChartType, hasPivotTableSource: Boolean, lastTag: Stack<String>): Series? {
            try {
                var eventType = xpp.eventType
                var idx = 0
                val seriesidx = parentChart.parentChart!!.allSeries.size
                val ranges = arrayOf("", "", "", "")     //legend, cat, ser/value, bubble cell references
                var legendText: String? = ""
                var sp: SpPr? = null
                var d: DLbls? = null
                var m: Marker? = null
                var smooth = false
                var dpts: ArrayList<*>? = null
                var cache: String? = null
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        var tnm = xpp.name
                        if (tnm == "ser") { // Represents a single field in the PivotTable. This complex type contains
                            idx = 0
                        } else if (tnm == "order") {// attr val= order
                        } else if (tnm == "cat" || tnm == "xVal") {        // children: CHOICE OF: multiLvlStrRef, numLiteral, numRef, strLit, strRef
                            idx = 1
                        } else if (tnm == "val" || tnm == "yVal") {        // children: CHOICE OF: numLit, numRef
                            idx = 2
                        } else if (tnm == "dLbls") {        // data labels
                            lastTag.push(tnm)                // keep track of element hierarchy
                            d = DLbls.parseOOXML(xpp, lastTag, wbh).cloneElement() as DLbls
                            if (d.showBubbleSize()) parentChart.setChartOption("ShowBubbleSizes", "1")
                            if (d.showCatName()) parentChart.setChartOption("ShowCatLabel", "1")
                            if (d.showLeaderLines()) parentChart.setChartOption("ShowLdrLines", "1")
                            if (d.showLegendKey())
                            ; // TODO: handle show legend key
                            if (d.showPercent()) parentChart.setChartOption("ShowLabelPct", "1")
                            if (d.showSerName()) parentChart.setChartOption("ShowLabel", "1")
                            if (d.showVal()) parentChart.setChartOption("ShowValueLabel", "1")
                            // data label options
                        } else if (tnm == "dPt") {            // data point(s)
                            if (dpts == null) dpts = ArrayList()
                            lastTag.push(tnm)                // keep track of element hierarchy
                            dpts!!.add(DPt.parseOOXML(xpp, lastTag, wbh).cloneElement())
                        } else if (tnm == "spPr") {    // series spPr
                            lastTag.push(tnm)                // keep track of element hierarchy
                            sp = SpPr.parseOOXML(xpp, lastTag, wbh).cloneElement() as SpPr
                        } else if (tnm == "marker") {
                            lastTag.push(tnm)                // keep track of element hierarchy
                            m = Marker.parseOOXML(xpp, lastTag, wbh).cloneElement() as Marker
                        } else if (tnm == "bubbleSize") {
                            idx = 3
                        } else if (tnm == "shape") {    // bar only
                            parentChart.convertShape(xpp.getAttributeValue(0))
                        } else if (tnm == "smooth") {  // line chart
                            smooth = xpp.attributeCount == 0 || xpp.getAttributeValue(0) != "0"
                        } else if (tnm == "explosion") {
                            val v = xpp.getAttributeValue(0)
                            parentChart.setChartOption("Percentage", v)
                            // NOTE: two types of values; 1- reference denoted by f element parents can be numRef, strRef or multiLvlStrRef
                            //						  2- text value denoted by v element parent is strRef
                        } else if (tnm == "formatCode") {    // part of numCache element
                            Xf.addFormatPattern(wbh.workBook!!, OOXMLAdapter.getNextText(xpp))
                            // *******************************************
                            // TODO: add to y pattern ***
                            // *******************************************
                        } else if (tnm == "f") {    // range element  -- legend cell, Cat range, Value range, Bubble data reference
                            ranges[idx] = OOXMLAdapter.getNextText(xpp)
                        } else if (tnm == "v") {    // value or text of series or category (parent=tx)
                            if (idx == 0)
                            // legend text
                                legendText = OOXMLAdapter.getNextText(xpp)    // legend text; possible to have legend text without a legend cell range (ranges[1])
                            else if (idx == -1 || ranges[idx] == "") { // shoudln't!! can't have a textual refernce in place of a series or cat value (can you?)
                                Logger.logWarn("ChartSeries.parseOOXML: unexpected text value")
                            }
                        } else if (tnm == "numCache" || tnm == "strCache" || tnm == "multiLvlStrRef") {    // parent= cat or vals (series values)
                            cache = tnm
                        } else if (tnm == "ptCount") {    // parent= numCache or strCache (governs either f element)
                            if (hasPivotTableSource) {    // OK, if have a pivot table source then the range referenced in f is only a SUBSET
                                // unclear if at any other time the range referenced is a subset ... [NOTE: in testing, only pivot charts hit]
                                // another assumption:  assume that range is only TRUNCATED -- in testing, true
                                val npoints = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                                if (ranges[idx] != "" && ranges[idx].indexOf(",") == -1) {
                                    try {
                                        val cells = CellRange(ranges[idx], wbh, false, true)
                                        if (cells.getCells()!!.size != npoints) {    //must adjust
                                            val z = 0
                                            val clist = cells.getCells()
                                            while (eventType != XmlPullParser.END_DOCUMENT) {
                                                if (eventType == XmlPullParser.START_TAG) {
                                                    tnm = xpp.name
                                                    if (tnm == "pt") {
                                                        // format code idx
                                                    } else if (tnm == "v") {
                                                        /* this case should NOT happen
													String s= OOXMLAdapter.getNextText(xpp);
 													if (z < clist.length)
														if (!clist[z].getVal().toString().equals(s))
															Logger.logWarn("ChartSeries.parseOOXML: unexpected pivot value order- skipping");
													z++;
*/
                                                    }
                                                } else if (eventType == XmlPullParser.END_TAG) {
                                                    if (xpp.name == cache) {
                                                        cache = null
                                                        break
                                                    }
                                                }
                                                eventType = xpp.next()
                                            }
                                            // pivot charts: apparently always truncate/skip last cell in range (which represents the grand total)
                                            if (npoints < clist!!.size) {// truncate!
                                                val rc = cells.rangeCoords
                                                rc[0]--
                                                rc[2]--    // make 0-based
                                                if (rc[0] == rc[2])
                                                    rc[3] -= clist.size - npoints
                                                else
                                                    rc[2] -= clist.size - npoints
                                                // KSC: TESTING: REMOVE WHEN DONE
                                                //io.starter.toolkit.Logger.log("Truncate list: old range: " + ranges[idx] + " new range: " + cells.getSheet().getQualifiedSheetName() + "!" + ExcelTools.formatLocation(rc));
                                                ranges[idx] = cells.sheet!!.qualifiedSheetName + "!" + ExcelTools.formatLocation(rc)
                                            }

                                            continue    // don't hit xpp.next() below
                                        }
                                    } catch (e: Exception) {
                                        Logger.logErr("ChartSeries.parseOOXML: Error adjusting pivot range for $parentChart:$e")
                                    }
                                    // problems parsing range - skp
                                }
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        if (xpp.name == "ser") {
                            lastTag.pop()
                            // must only have 1 for pie-type charts (pie, of pie - bar of, pie of)
                            val s = parentChart.parentChart!!.chartSeries.addSeries(ranges[2], ranges[1], ranges[3], ranges[0], legendText, parentChart, parentChart.parentChart!!.getChartOrder(parentChart))
                            if (sp != null)
                                s.spPr = sp
                            else if (seriesidx < 7)
                            // TODO: figure out where to get colors past 6
                                if (seriesidx > 0 && parentChart is PieChart)
                                    Logger.logWarn("ChartSeries.parseOOXML:  more than 1 series encountered for a Pie-style chart")
                                else
                                    s.setColor(wbh.workBook!!.theme!!.genericThemeClrs[seriesidx + 4])// TODO: " When you create a chart, by default - the first six series are the six accent colors in order - but not the exact color or any variation that appears in the palette. They're typically (unless the primary accent color being modified is extremely dark) a bit darker than the primary accent color. Chart series 7 - 12 use the actual primary accent colors 1 through 6 ... and then chart series 13 starts a set of lighter variations of the six accent colors that are also slightly different from any position in the palette."
                            // series colors start at 4
                            if (d != null)
                                s.dLbls = d
                            if (dpts != null) {
                                for (z in dpts.indices)
                                    s.addDpt(dpts[z] as DPt)
                            }
                            if (m != null)
                                s.marker = m
                            if (smooth)
                                s.setHasSmoothLines(smooth)
                            return s
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("ChartSeries.parseOOXML: Error parsing series for $parentChart:$e")
            }

            return null
        }
    }


}
