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

import io.starter.formats.XLS.WorkBook
import io.starter.toolkit.Logger

import java.util.ArrayList
import java.util.HashMap

open class ColChart(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) : ChartType(charttype, cf, wb) {
    protected var col: Bar? = null

    /**
     * @return truth of "Chart is Clustered"  (Bar/Col only)
     */
    override/*cf.isClustered());	*/ val isClustered: Boolean
        get() = !isStacked && !is100PercentStacked

    init {
        col = charttype as Bar
        defaultShape = ChartConstants.SHAPEDEFAULT
    }

    /**
     * returns SVG to represent the actual chart object i.e. the representation of the series data in the particular format (BAR, LINE, AREA, etc.)
     *
     * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
     * @param axisMetrics  maps specific axis options such as xAxisReversed, xPattern ...
     * @param categories
     * @param series       arraylist of double[] series values
     * @param seriescolors int[] of series or bar colors color ints
     * @return String svg
     */
    override fun getSVG(chartMetrics: HashMap<String, Double>, axisMetrics: HashMap<String, Any>, s: ChartSeries): String {
        val x = chartMetrics["x"]    // + (!yAxisReversed?0:w);	// x is constant at x origin unless reversed
        val y = chartMetrics["y"]
        val w = chartMetrics["w"]
        val h = chartMetrics["h"]
        val max = chartMetrics["max"]
        val min = chartMetrics["min"]
        val categories = s.getCategories()
        val series = s.seriesValues
        val seriescolors = s.seriesBarColors
        val legends = s.getLegends()
        if (series!!.size == 0) {
            Logger.logErr("Bar.getSVG: error in series")
            return ""
        }
        val n = series.size
        val dls = dataLabelInts // get array of data labels (can be specific per series ...)
        val isXReversed = axisMetrics["xAxisReversed"] as Boolean
        val isYReversed = axisMetrics["yAxisReversed"] as Boolean

        val svg = StringBuffer()        // return svg
        svg.append("<g>\r\n")

        var barw = 0.0
        var yfactor = 0.0    //
        if (n > 0) {
            barw = w / (categories!!.size * (n + 1.0) + 1)    // w/(#cats * nseries+1) + 1
            if (max != 0.0)
                yfactor = h / max    // h/YMAXSCALE
        }
        val rfX = if (!isXReversed) 1 else -1 // reverse factor
        val rfY = if (!isYReversed) 1 else -1
        // for each series
        for (i in 0 until n) {    // each series group
            svg.append("<g>\r\n")
            val y0 = y + if (!isYReversed) h else 0    // start from bottom and work up (unless reversed)
            val curseries = series[i] as DoubleArray
            val curranges = s.seriesRanges!![i] as Array<String>

            //x+=barw;	// a barwidth separates each series group
            for (j in curseries.indices) {        // each series
                val xx = x + barw * (i + 1) + j.toDouble() * (n + 1).toDouble() * barw        // x goes from 1 series to next, corresponding to bar/column color
                val hh = yfactor * curseries[j]                        // bar height = measure of series value
                val yy = y0 - if (!isYReversed) hh else 0                // start drawing column
                svg.append("<rect id='series_" + (i + 1) + "' " + ChartType.getScript(curranges[j]) + " fill='" + seriescolors!![i] + "' fill-opacity='1' stroke='black' stroke-opacity='1' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'" +
                        " x='" + xx + "' y='" + yy + "' width='" + barw + "' height='" + hh + "' fill-rule='evenodd'/>")
                val l = getSVGDataLabels(dls, axisMetrics, curseries[j], 0.0, i, legends, categories!![j].toString())
                if (l != null) {
                    svg.append("<text x='" + (xx + barw / 2) + "' y='" + (y0 - (hh + 10) * rfY) +
                            "' style='text-anchor: middle;' " + dataLabelFontSVG + ">" + l + "</text>\r\n")
                }
            }
            svg.append("</g>\r\n")        // each series group
        }
        svg.append("</g>\r\n")
        return svg.toString()
    }

    /**
     * gets the chart-type specific ooxml representation: <colChart>
     *
     * @return
    </colChart> */
    override fun getOOXML(catAxisId: String, valAxisId: String, serAxisId: String): StringBuffer? {
        val cooxml = StringBuffer()

        // chart type: contains chart options and series data
        cooxml.append("<c:barChart>")
        cooxml.append("\r\n")
        cooxml.append("<c:barDir val=\"col\"/>")
        cooxml.append("<c:grouping val=\"")


        if (this.is100PercentStacked)
            cooxml.append("percentStacked")
        else if (this.isStacked)
            cooxml.append("stacked")
        else if (this.isClustered)
            cooxml.append("clustered")
        else
            cooxml.append("standard")
        cooxml.append("\"/>")
        cooxml.append("\r\n")
        // vary colors???

        // *** Series Data:	ser, cat, val for most chart types
        cooxml.append(this.parentChart!!.chartSeries.getOOXML(this.chartType, false, 0))

        // chart data labels, if any
        //TODO: FINISH
        //cooxml.append(getDataLabelsOOXML(cf));
        //TODO: FINISH DR0P LINES
        ///    			if (this.hasDropLines() )	/* Area, Line, Stock*/
        //				cooxml.append()
        if (this.getChartOption("Gap") != "150")
            cooxml.append("<c:gapWidth val=\"" + this.getChartOption("Gap") + "\"/>")    // default= 0
        if (this.getChartOption("Overlap") != "0")
            cooxml.append("<c:overlap val=\"" + this.getChartOption("Overlap") + "\"/>")    // default= 0
        // Series Lines
        val cl = cf!!.chartLinesRec
        if (cl != null)
            cooxml.append(cl.ooxml)

        // axis ids	 - unsigned int strings
        cooxml.append("<c:axId val=\"$catAxisId\"/>")
        cooxml.append("\r\n")
        cooxml.append("<c:axId val=\"$valAxisId\"/>")
        cooxml.append("\r\n")

        cooxml.append("</c:barChart>")
        cooxml.append("\r\n")

        return cooxml
    }
}
