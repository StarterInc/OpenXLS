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

class StackedAreaChart(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) : AreaChart(charttype, cf, wb) {

    override var isStacked: Boolean
        get() = true
        set(value: Boolean) {
            super.isStacked = value
        }

    init {
        area = charttype as Area
    }

    /**
     * returns SVG to represent the actual chart object (BAR, LINE, AREA, etc.)
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
        if (series!!.size == 0) {
            Logger.logErr("Area.getSVG: error in series")
            return ""
        }
        val svg = StringBuffer()
        val dls = dataLabelInts // get array of data labels (can be specific per series ...)

        var xfactor = 0.0
        var yfactor = 0.0    //
        if (categories!!.size > 1)
            xfactor = w / (categories.size - 1)    // w/#categories
        if (max != 0.0)
            yfactor = h / max    // h/YMAXSCALE

        // first, calculate Area points- which are summed per series
        val n = series.size
        val nSeries = (series[0] as DoubleArray).size
        val areapoints = Array(nSeries) { DoubleArray(n) }
        for (i in 0 until n) {
            val seriesy = series[i] as DoubleArray
            for (j in seriesy.indices) {
                val yval = seriesy[j]
                areapoints[j][i] = yval + if (i > 0) areapoints[j][i - 1] else 0
            }
        }
        // for each series
        for (i in n - 1 downTo 0) {    // "paint" right to left
            svg.append("<g>\r\n")
            var points = ""
            var x1 = 0.0
            var y1 = 0.0
            var labels: String? = null
            val curseries = series[i] as DoubleArray
            for (j in curseries.indices) {
                x1 = x + j * xfactor
                val yval = areapoints[j][i]    // current point
                points += x + j * xfactor + "," + (y + h - yval * yfactor)

                if (j == 0) y1 = y + h - yval * yfactor    // end point (==start point) for path statement below
                points += " "
                // DATA LABELS
                val l = getSVGDataLabels(dls, axisMetrics, curseries[j], 0.0, i, legends, categories[j].toString())
                if (l != null) {
                    // if only category label, center over all series; anything else, position at data point
                    val showCategories = dls[i] and AttachedLabel.CATEGORYLABEL == AttachedLabel.CATEGORYLABEL
                    val showValueLabel = dls[i] and AttachedLabel.VALUELABEL == AttachedLabel.VALUELABEL
                    val showValue = dls[i] and AttachedLabel.VALUE == AttachedLabel.VALUE
                    if (showCategories && !(showValue || showValueLabel) && j == 0) {    // only 1 label, centered along category axis within area
                        //y1+= (seriesx[seriesx.length/2]/2)*yfactor;
                        val hh = areapoints[areapoints.size / 2][i] * yfactor
                        val yy = y + h - hh + 10
                        if (labels == null) labels = ""
                        labels = "<text x='" + (x + w / 2) + "' y='" + yy + "' vertical-align='middle' " + dataLabelFontSVG + " style='text-align:middle;'>" + l + "</text>\r\n"
                    } else if (showValue || showValueLabel) { // labels at each data point
                        if (labels == null) labels = ""
                        val yy = y + h - (yval - curseries[j] * .5) * yfactor
                        labels += "<text x='$x1' y='$yy' style='text-anchor: middle;' $dataLabelFontSVG>$l</text>\r\n"
                    }
                }
            }
            // pointsends connects up area to beginning
            val pointsend = x1.toString() + "," + (y + h) +
                    " " + x + "," + (y + h) +
                    " " + x + "," + y1
            svg.append("<polyline  id='series_" + (i + 1) + "' " + ChartType.getScript("") + " fill='" + seriescolors!![i] + "' fill-opacity='1' " + strokeSVG + " points='" + points + pointsend + "' fill-rule='evenodd'/>\r\n")

            /* john took out

			// do twice to make slightly thicker
			svg.append("<polyline fill='none' fill-opacity='0' stroke='black' stroke-opacity='1' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'" +
					" points='" + points + "'/>\r\n");
			// do twice to make slightly thicker
			svg.append("<polyline fill='none' fill-opacity='0' stroke='black' stroke-opacity='1' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'" +
					" points='" + points + "'/>\r\n");*/
            // Now print data labels, if any
            if (labels != null) svg.append(labels)
            svg.append("</g>\r\n")
        }
        return svg.toString()
    }
}
