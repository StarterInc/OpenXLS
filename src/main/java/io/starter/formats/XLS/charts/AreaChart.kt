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
package io.starter.formats.XLS.charts

import io.starter.OpenXLS.CellRange
import io.starter.OpenXLS.ChartSeriesHandle
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.XLS.FormatConstants
import io.starter.formats.XLS.WorkBook
import io.starter.toolkit.Logger
import org.json.JSONArray
import org.json.JSONException
import org.json.JSONObject

import java.util.ArrayList
import java.util.HashMap

/**
 * non-stacked area chart
 */
open class AreaChart(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) : ChartType(charttype, cf, wb) {
    internal var area: Area? = null

    /**
     * return the type JSON for this Chart Object
     *
     * @return
     */
    override val typeJSON: JSONObject
        @Throws(JSONException::class)
        get() {
            val typeJSON = JSONObject()
            val dojoType: String
            if (!this.isStacked) {
                dojoType = "Areas"
            } else {
                dojoType = "StackedAreas"
            }
            typeJSON.put("type", dojoType)
            return typeJSON
        }

    init {
        area = charttype as Area
    }

    @Throws(JSONException::class)
    fun getJSON(series: Array<ChartSeriesHandle>, wbh: WorkBookHandle, minMax: Array<Double>): JSONObject {
        val chartObjectJSON = JSONObject()
        // Type JSON
        chartObjectJSON.put("type", this.typeJSON)

        // Series
        var yMax = 0.0
        var yMin = 0.0
        var nSeries = 0
        val seriesJSON = JSONArray()
        val seriesCOLORS = JSONArray()
        try {
            for (i in series.indices) {
                val seriesvals = CellRange
                        .getValuesAsJSON(series[i].seriesRange, wbh)
                // must trap min and max for axis tick and units
                var sum = 0.0 // for area-type charts, ymax is the sum of
                // all points in same series
                nSeries = Math.max(nSeries, seriesvals.length())
                for (j in 0 until seriesvals.length()) {
                    try {
                        sum += seriesvals.getDouble(j)
                        yMax = Math.max(yMax, sum)
                        yMin = Math.min(yMin, seriesvals.getDouble(j))
                    } catch (n: NumberFormatException) {
                    }

                }
                seriesJSON.put(seriesvals)
                seriesCOLORS.put(FormatConstants.SVGCOLORSTRINGS[series[i]
                        .seriesColor])
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
        chartMetrics["min"]
        val categories = s.getCategories()
        val series = s.seriesValues
        val seriescolors = s.seriesBarColors
        val legends = s.getLegends()
        // x value for each point= w/(ncategories + 1) 1st one is
        // xv*2 then increases from there
        // y value for each point= h/YMAX
        if (series!!.size == 0) {
            Logger.logErr("Area.getSVG: error in series")
            return ""
        }
        val svg = StringBuffer()
        val dls = dataLabelInts // get array of data labels (can
        // be specific per series ...)

        var xfactor = 0.0
        var yfactor = 0.0 //
        if (categories!!.size > 1) {
            xfactor = w / (categories.size - 1) // w/#categories
        }
        if (max != 0.0) {
            yfactor = h / max // h/YMAXSCALE
        }

        // for each series
        val n = series.size
        for (i in n - 1 downTo 0) { // "paint" right to left
            svg.append("<g>\r\n")
            var points = ""
            var x1 = 0.0
            var y1 = 0.0
            var labels: String? = null
            val curseries = series[i] as DoubleArray
            for (j in curseries.indices) {
                x1 = x + j * xfactor
                val yval = curseries[j] // areapoints[j][i]; //
                // current point
                points += (x + j * xfactor + ","
                        + (y + h - yval * yfactor))

                if (j == 0) {
                    y1 = y + h - yval * yfactor // end point (==start
                    // point) for path
                    // statement below
                }
                points += " "
                // DATA LABELS
                val l = getSVGDataLabels(dls, axisMetrics, curseries[j], 0.0, i, legends, categories[j]
                        .toString())
                if (l != null) {
                    // if only category label, center over all series; anything
                    // else, position at data point
                    val showCategories = dls[i] and AttachedLabel.CATEGORYLABEL == AttachedLabel.CATEGORYLABEL
                    val showValueLabel = dls[i] and AttachedLabel.VALUELABEL == AttachedLabel.VALUELABEL
                    val showValue = dls[i] and AttachedLabel.VALUE == AttachedLabel.VALUE
                    if (showCategories && !(showValue || showValueLabel)
                            && j == 0) { // only 1 label, centered along
                        // category axis within area
                        val hh = y1 // (areapoints[areapoints.length/2][i]*yfactor);
                        val yy = y + h - hh + 10
                        if (labels == null) {
                            labels = ""
                        }
                        labels = ("<text x='" + (x + w / 2) + "' y='" + yy
                                + "' vertical-align='middle' "
                                + ChartType.dataLabelFontSVG
                                + " style='text-align:middle;'>" + l
                                + "</text>\r\n")
                    } else if (showValue || showValueLabel) { // labels at each
                        // data point
                        if (labels == null) {
                            labels = ""
                        }
                        val yy = y + h - (yval - curseries[j] * .5) * yfactor
                        labels += ("<text x='" + x1 + "' y='" + yy
                                + "' style='text-anchor: middle;' "
                                + ChartType.dataLabelFontSVG/*
                         * +" fill='"
                         * +getDarkColor
                         * ()+"'
                         */ + ">"
                                + l + "</text>\r\n")
                    }
                }
            }
            // pointsends connects up area to beginning
            val pointsend = (x1.toString() + "," + (y + h) + " " + x + ","
                    + (y + h) + " " + x + "," + y1)
            // String clr= getDarkColor();
            /*
             * try { clr=
             * FormatConstants.SVGCOLORSTRINGS[seriescolors[i]]; }
             * catch(ArrayIndexOutOfBoundsException e) {; }
             */
            svg.append("<polyline  id='series_" + (i + 1) + "' " + ChartType.getScript("")
                    + " fill='" + seriescolors!![i] + "' fill-opacity='1' "
                    + strokeSVG + " points='" + points + pointsend
                    + "' fill-rule='evenodd'/>\r\n")

            // Now print data labels, if any
            if (labels != null) {
                svg.append(labels)
            }
            svg.append("</g>\r\n")
        }
        return svg.toString()
    }

    /**
     * gets the chart-type specific ooxml representation: <areaChart>
     *
     * @return
    </areaChart> */
    override fun getOOXML(catAxisId: String, valAxisId: String, serAxisId: String): StringBuffer? {
        val cooxml = StringBuffer()

        // chart type: contains chart options and series data
        cooxml.append("<c:areaChart>")
        cooxml.append("\r\n")
        cooxml.append("<c:grouping val=\"")

        if (this.is100PercentStacked) {
            cooxml.append("percentStacked")
        } else if (this.isStacked) {
            cooxml.append("stacked")
            // } else if (this.isClustered())
            // grouping="clustered";
        } else {
            cooxml.append("standard")
        }
        cooxml.append("\"/>")
        cooxml.append("\r\n")
        // vary colors???

        // *** Series Data: ser, cat, val for most chart types
        cooxml.append(this.parentChart!!.chartSeries
                .getOOXML(this.chartType, false, 0))

        // TODO: FINISH
        // chart data labels, if any
        // cooxml.append(getDataLabelsOOXML(cf));
        if (this.cf!!.chartLinesRec != null) {
            cooxml.append(this.cf!!.chartLinesRec.ooxml)
        }

        // axis ids - unsigned int strings
        cooxml.append("<c:axId val=\"$catAxisId\"/>")
        cooxml.append("\r\n")
        cooxml.append("<c:axId val=\"$valAxisId\"/>")
        cooxml.append("\r\n")

        cooxml.append("</c:areaChart>")
        cooxml.append("\r\n")

        return cooxml
    }

}
