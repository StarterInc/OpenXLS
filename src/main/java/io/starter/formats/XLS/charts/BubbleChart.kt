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

class BubbleChart(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) : ChartType(charttype, cf, wb) {

    private var bubble: Scatter? = null
    /**
     * return true if this bubble chart series should be displayed as 3d
     *
     * @return
     */
    /**
     * Bubble charts handle 3d differently
     *
     * @param is3d
     */
    var is3d = false

    init {
        bubble = charttype as Scatter
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
        val x = chartMetrics["x"]
        val y = chartMetrics["y"]
        val w = chartMetrics["w"]
        val h = chartMetrics["h"]
        val max = chartMetrics["max"]
        val min = chartMetrics["min"]
        val categories = s.getCategories()
        val series = s.seriesValues
        val seriescolors = s.seriesBarColors
        val legends = s.getLegends()
        // x value for each point= w/(ncategories + 1) 1st one is xv*2 then increases from there
        // y value for each point= h/YMAX
        val svg = StringBuffer()

        if (series!!.size == 0) {
            Logger.logErr("Scatter.getSVG: error in series")
            return ""
        }
        // gather data labels, markers, has lines, series colors for chart
        val threeD = cf!!.isThreeD(ChartConstants.BUBBLECHART)
        val dls = dataLabelInts // get array of data labels (can be specific per series ...)
        val hasLines = cf!!.hasLines // should be per-series?
        val markers = markerFormats    // get an array of marker formats per series
        val n = series.size
        var seriesx: DoubleArray? = null
        var xfactor = 0.0
        var yfactor = 0.0
        var bfactor = 0.0    //
        var TEXTUALXAXIS = true
        // get x axis max/min for an x axis which is a value axis
        seriesx = DoubleArray(categories!!.size)
        var xmin = java.lang.Double.MAX_VALUE
        var xmax = java.lang.Double.MIN_VALUE
        for (j in categories.indices) {
            try {
                seriesx[j] = Double(categories[j].toString())
                xmax = Math.max(xmax, seriesx[j])
                xmin = Math.min(xmin, seriesx[j])
                TEXTUALXAXIS = false        // if ANY category val is a double, assume it's a normal xyvalue axis
            } catch (e: Exception) {
                /* keep going */
            }

        }
        if (!TEXTUALXAXIS) {
            val d = ValueRange.calcMaxMin(xmax, xmin, w)
            xfactor = w / d[2]        // w/maximum scale
        } else
            xfactor = w / (categories.size + 1)    // w/#categories

        if (max - min != 0.0)
            yfactor = h / Math.abs(max - min)    // h/YMAXSCALE
        // get Bubble Size Factor
        for (i in 0 until n) {
            val seriesy = series[i] as DoubleArray
            val curranges = s.seriesRanges!![i] as Array<String>
            val nseries = seriesy.size / 2
            var bmin = java.lang.Double.MAX_VALUE
            var bmax = java.lang.Double.MIN_VALUE
            for (j in nseries until seriesy.size) {
                bmax = Math.max(bmax, seriesy[j])
                bmin = Math.min(bmin, seriesy[j])
            }
            if (bmax - bmin != 0.0)
                bfactor = h / Math.abs(bmax - bmin) / 5.0
        }
        svg.append("<g>\r\n")
        if (threeD) {
            svg.append(get3DBubbleSVG(seriescolors!!))
        }
        // for each series
        for (i in 0 until n) {
            var labels = ""
            val seriesy = series[i] as DoubleArray
            val curranges = s.seriesRanges!![i] as Array<String>
            val nseries = seriesy.size / 2
            for (j in 0 until nseries) {
                var xval = 0.0
                if (TEXTUALXAXIS /*|| i > 0*/)
                    xval = (j + 1).toDouble()
                else
                    xval = seriesx[j]
                val cx = x + xval * xfactor
                val cy = y + h - seriesy[j] * yfactor
                val r = seriesy[j + nseries] * bfactor
                //io.starter.toolkit.Logger.log("x: " + xval + " val: " + seriesy[j] + " size: " + seriesy[j+nseries]);
                if (!threeD)
                    svg.append("<circle " + ChartType.getScript(curranges[j]) + " cx='" + cx + "' cy='" + cy + "' r='" + r + "' " + strokeSVG + " fill='" + seriescolors!![i] + "'/>\r\n")
                else
                    svg.append("<circle " + ChartType.getScript(curranges[j]) + "   id='series_" + (i + 1) + "' cx='" + cx + "' cy='" + cy + "' r='" + r + "' " + strokeSVG + " style='fill:url(#fill" + i + ")'/>\r\n")
                val l = getSVGDataLabels(dls, axisMetrics, seriesy[j + nseries], 0.0, i, legends, categories[j].toString())
                if (l != null)
                    labels += "<text x='" + (r + 10.0 + x + xval * xfactor) + "' y='" + (y + h - seriesy[j] * yfactor) + "' " + dataLabelFontSVG + ">" + l + "</text>\r\n"
            }
            // labels after lines and markers
            svg.append(labels)
        }
        //io.starter.toolkit.Logger.log("Bubble svg: " + svg.toString());
        svg.append("</g>\r\n")
        return svg.toString()
    }

    /**
     * returns the SVG necessary to define 3D bubbles (circles) for each series color
     *
     * @param seriescolors
     * @return String SVG
     */
    private fun get3DBubbleSVG(seriescolors: Array<String>): String {
        val svg = StringBuffer()
        svg.append("<defs>\r\n")
        for (i in seriescolors.indices) {
            svg.append("<radialGradient id='fill" + i + "' " +
                    "gradientUnits=\"objectBoundingBox\" fx=\"40%\" fy=\"30%\">")
            svg.append("<stop offset='0%' style='stop-color:#FFFFFF' />")
            svg.append("<stop offset='40%' style='stop-color:" + seriescolors[i] + "' stop-opacity='.65' />")
            svg.append("<stop offset='95%' style='stop-color:" + seriescolors[i] + "' stop-opacity='1' />")
            svg.append("<stop offset='99%' style='stop-color:" + seriescolors[i] + "' stop-opacity='.3' />")
            svg.append("<stop offset='100%' style='stop-color:" + seriescolors[i] + "'/>")
            svg.append("</radialGradient>\r\n")
        }
        svg.append("</defs>\r\n")
        return svg.toString()
    }

    /**
     * gets the chart-type specific ooxml representation: <bubbleChart>
     *
     * @return
    </bubbleChart> */
    override fun getOOXML(catAxisId: String, valAxisId: String, serAxisId: String): StringBuffer? {
        val cooxml = StringBuffer()

        // chart type: contains chart options and series data
        cooxml.append("<c:bubbleChart>")
        cooxml.append("\r\n")
        // vary colors???

        // *** Series Data:	ser, cat, val for most chart types
        cooxml.append(this.parentChart!!.chartSeries.getOOXML(this.chartType, false, 0))

        // chart data labels, if any
        //TODO: FINISH
        //cooxml.append(getDataLabelsOOXML(cf));
        if (this.is3d)
            cooxml.append("bubble3d val=\"1\"")
        // bubblescale
        cooxml.append("<c:bubbleScale val=\"100\"/>")    // TODO: read correct value
        // showNegBubbles
        // sizeRepresents

        // axis ids	 - unsigned int strings
        cooxml.append("<c:axId val=\"$catAxisId\"/>")
        cooxml.append("\r\n")
        cooxml.append("<c:axId val=\"$valAxisId\"/>")
        cooxml.append("\r\n")

        cooxml.append("</c:bubbleChart>")
        cooxml.append("\r\n")

        return cooxml
    }

}