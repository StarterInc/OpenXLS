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

class StackedColumn(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) : ColChart(charttype, cf, wb) {
    internal var defaultShape = 0 //????

    override var isStacked: Boolean
        get() = true
        set(value: Boolean) {
            super.isStacked = value
        }

    init {
        col = charttype as Bar
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
        val categories = s.getCategories()
        val series = s.seriesValues
        val seriescolors = s.seriesBarColors
        val legends = s.getLegends()
        if (series!!.size == 0) {
            Logger.logErr("Bar.getSVG: error in series")
            return ""
        }
        /*
         * TODO: Stacked vs 100% Stacked
         */
        val svg = StringBuffer()
        val dls = dataLabelInts // get array of data labels (can be specific per series ...)

        var barw = 0.0
        var yfactor = 0.0
        if (categories!!.size > 0)
            barw = w / (categories.size * 2)    // w/#categories (only 1 column per series)
        if (max != 0.0)
            yfactor = h / max    // h/YMAXSCALE

        var totalperseries: DoubleArray? = null
        val f100 = col!!.is100Percent
        if (f100) {  //FOR 100% STACKED
            // first, calculate points- which are summed per series
            val n = series.size
            val nSeries = (series[0] as DoubleArray).size
            totalperseries = DoubleArray(nSeries)
            for (i in 0 until n) {
                val seriesy = series[i] as DoubleArray
                for (j in seriesy.indices) {
                    val yval = seriesy[j]
                    totalperseries[j] = yval + totalperseries[j]
                }
            }
        }
        // for each series - ONE COLUMN per series
        var previousY = DoubleArray(0)
        for (i in series.indices) {
            svg.append("<g>\r\n")
            val curseries = series[i] as DoubleArray    // for each data point - stacked on series column
            val curranges = s.seriesRanges!![i] as Array<String>
            var xx: Double
            var yy = y + h    // origin
            if (i == 0) {
                previousY = DoubleArray(curseries.size)
                for (j in previousY.indices)
                    previousY[j] = yy    // origin
            }
            var barh: Double
            for (j in curseries.indices) {
                xx = x + j.toDouble() * barw * 2.0 + barw / 2
                if (previousY.size > j)
                // should
                    yy = previousY[j]
                if (f100)
                    barh = curseries[j] / totalperseries!![j] * (h - 10)    // height of current point as a percentage of total points per series
                else
                    barh = curseries[j] * yfactor    // height of current point
                svg.append("<rect " + ChartType.getScript(curranges[j]) + " fill='" + seriescolors!![i] + "' fill-opacity='1' " + strokeSVG +
                        " x='" + xx + "' y='" + (yy - barh) + "' width='" + barw + "' height='" + barh + "'/>")
                //TODO: DATA LABELS
                // Now print data labels, if any
                //if (labels!=null)  svg.append(labels);
                if (previousY.size > j)
                // should
                    previousY[j] = yy - barh
            }
            svg.append("</g>\r\n")
        }
        return svg.toString()
    }
}
