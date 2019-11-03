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

class DoughnutChart(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) : ChartType(charttype, cf, wb) {
    private var doughnut: Pie? = null

    /**
     * return the type JSON for this Chart Object
     *
     * @return
     */
    override// default
    // default - rounds percentages up
    // TODO: Interpret distance ...
    val typeJSON: JSONObject
        @Throws(JSONException::class)
        get() {
            val typeJSON = JSONObject()
            typeJSON.put("labelOffset", -25)
            typeJSON.put("precision", 0)
            typeJSON.put("type", "Pie")
            return typeJSON
        }

    init {
        doughnut = charttype as Pie
    }

    /**
     * return the JSON that
     *
     * @param seriesvals
     * @param range
     * @return
     */
    @Throws(JSONException::class)
    fun getJSON(series: Array<ChartSeriesHandle>, wbh: WorkBookHandle, minMax: Array<Double>): JSONObject {
        val chartObjectJSON = JSONObject()
        // Type JSON
        chartObjectJSON.put("type", this.typeJSON)

        // TODO: check out labels options chosen: default label is Category name + percentage ...
        var yMax = 0.0
        var yMin = 0.0
        var len = 0.0
        val pieSeries = JSONArray()
        try {
            val range = series[0].categoryRange        // 20080516 KSC: retrieve cat range instead of parameter
            val cats = CellRange.getValuesAsJSON(range, wbh)    // parse category range into JSON Array
            val seriesvals = CellRange.getValuesAsJSON(series[0].seriesRange, wbh)
            var piesum = 0.0
            for (k in 0 until seriesvals.length()) {
                piesum += seriesvals.getDouble(k)
                yMax = Math.max(yMax, seriesvals.getDouble(k))
                yMin = Math.min(yMin, seriesvals.getDouble(k))
            }
            val percent = 100 / piesum
            for (k in 0 until seriesvals.length()) {
                val piepoint = JSONObject()
                piepoint.put("y", seriesvals.getDouble(k))
                piepoint.put("text", cats.getString(k) + "\n" + Math.round(percent * seriesvals.getDouble(k)) + "%")
                piepoint.put("color", FormatConstants.SVGCOLORSTRINGS[series[k].getPieChartSliceColor(k)])
                piepoint.put("stroke", darkColor)
                pieSeries.put(piepoint)
            }
            len = seriesvals.length().toDouble()
        } catch (e: Exception) {
            // TODO: warning ...?
        }

        // 20090717 KSC: input outside of try/catch to always set
        minMax[0] = yMin
        minMax[1] = yMax
        minMax[2] = len
        chartObjectJSON.put("Series", pieSeries)
        chartObjectJSON.put("SeriesFills", "")    // not applicable for pie charts; color is set above
        return chartObjectJSON
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
        val w = chartMetrics["w"]
        val h = chartMetrics["h"]
        val max = chartMetrics["max"]
        val min = chartMetrics["min"]
        val categories = s.getCategories()
        val series = s.seriesValues
        val seriescolors = s.seriesBarColors
        val legends = s.getLegends()
        val svg = StringBuffer()
        if (series!!.size == 0) {
            Logger.logErr("DoughnutChart.getSVG: error in series")
            return ""
        }
        var dls = dataLabelInts // get array of data labels (can be specific per series ...)
        val threeD = this.isThreeD
        val LABELOFFSET = 15

        val n = series.size
        var centerx = 0.0
        var centery = 0.0
        var radius = 0.0
        val radiusy = 0.0
        centerx = w / 2 + chartMetrics["x"]
        centery = h / 2 + chartMetrics["y"]
        if (threeD) {
            svg.append("<defs>\r\n")
            svg.append("<filter id=\"multiply\">\r\n")
            svg.append("<feBlend mode=\"multiply\" in2=\"\"/>\r\n")
            svg.append("</filter>\r\n")
            svg.append("</defs>\r\n")
        }
        svg.append("<g>\r\n")

        //radius= Math.min(w, h)/2.3;	// should take up almost entire w/h of chart
        radius = Math.min(w, h) / 1.9    // should take up almost entire w/h of chart
        val r0 = radius / 2 / n    // radius/2 to account for hole
        var r = radius    // start at outside and work inside
        for (i in n - 1 downTo 0) {    // for each series
            val curseries = series[i] as DoubleArray        // series data points
            var total = 0.0
            for (j in curseries.indices) {    // get total in order to calculate percentages
                total += curseries[j]
            }
            if (dls.size == 1) {    // no seriess-specific data labels; expand to entire series for loop below
                val dl = dls[0]
                dls = IntArray(curseries.size)
                for (z in dls.indices)
                    dls[z] = dl
            }
            svg.append("<circle " + ChartType.getScript("") + " cx='" + centerx + "' cy='" + centery + "' r='" + r + "' " + ChartType.getStrokeSVG(2f, darkColor) + " fill='none'/>\r\n")
            var x = centerx + r
            var y = centery
            var path = ""
            var percentage = 0.0
            var lasta = 0.0
            var largearcflag = 0
            var sweepflag = 0
            // Now create each pie wedge according to it's percentage value
            for (j in curseries.indices) {
                percentage = curseries[j] / total
                val angle = percentage * 360 + lasta
                val x1 = centerx + r * Math.cos(Math.toRadians(angle))
                val y1 = centery - r * Math.sin(Math.toRadians(angle))
                if (percentage * 360 > 180) {
                    sweepflag = 0
                    largearcflag = 1
                } else
                    largearcflag = 0
                path = "M" + centerx + " " + centery + " L" + x + " " + y + " A" + r + " " + r + " 0 " + largearcflag + " " + sweepflag + " " + x1 + " " + y1 + " L" + centerx + " " + centery + "Z"
                // paint wedge of color to center of chart -- the inner will overwrite so don't have to worry about segments and arcs, etc
                svg.append("<path  " + ChartType.getScript("") + "  fill='" + seriescolors!![j] + "'   id='series_" + (j + 1) + "' fill-opacity='" + fillOpacity + "' " + strokeSVG + " path='' d='" + path + "' fill-rule='evenodd'/>\r\n")
                // data labels
                val l = getSVGDataLabels(dls, axisMetrics, curseries[j], percentage, j, legends, categories!![j].toString())
                if (l != null) {
                    val halfa = percentage / 2 * 360 + lasta    // center in area
                    val x2 = centerx + (r - r0 / 2) * Math.cos(Math.toRadians(halfa))
                    val y2 = centery - (r - r0 / 2) * Math.sin(Math.toRadians(halfa))
                    svg.append("<text x='$x2' y='$y2' vertical-align='bottom' $dataLabelFontSVG style='text-anchor: middle;'>$l</text>\r\n")
                }
                lasta = angle
                x = x1
                y = y1
            }    // each point in current series
            r -= r0
        }    // each series
        // complete inner circle & create "hole"
        svg.append("<circle " + ChartType.getScript("") + " cx='" + centerx + "' cy='" + centery + "' r='" + r + "' " + ChartType.getStrokeSVG(2f, darkColor) + " fill='white'/>\r\n")
        svg.append("</g>\r\n")

        return svg.toString()
    }
    /**
     * Of the four candidate arc sweeps, two will represent an arc sweep of greater than or equal to 180 degrees (the "large-arc"),
     * and two will represent an arc sweep of less than or equal to 180 degrees (the "small-arc").
     * If large-arc-flag is '1', then one of the two larger arc sweeps will be chosen; otherwise, if large-arc-flag is '0', one of the smaller arc sweeps will be chosen,
     * If sweep-flag is '1', then the arc will be drawn in a "positive-angle" direction (i.e., the ellipse formula x=cx+rx*cos(theta)
     * and y=cy+ry*sin(theta) is evaluated such that theta starts at an angle corresponding to the current point and increases positively until the arc reaches (x,y)).
     * A value of 0 causes the arc to be drawn in a "negative-angle" direction
     * (i.e., theta starts at an angle value corresponding to the current point and decreases until the arc reaches (x,y)).
     */

    /**
     * gets the chart-type specific ooxml representation: <doughnutChart>
     *
     * @return
    </doughnutChart> */
    override fun getOOXML(catAxisId: String, valAxisId: String, serAxisId: String): StringBuffer? {
        val cooxml = StringBuffer()

        // chart type: contains chart options and series data
        cooxml.append("<c:doughnutChart>")
        cooxml.append("\r\n")
        // vary colors???
        cooxml.append("<c:varyColors val=\"1\"/>")

        // *** Series Data:	ser, cat, val for most chart types
        cooxml.append(this.parentChart!!.chartSeries.getOOXML(this.chartType, false, 0))

        // chart data labels, if any
        //TODO: FINISH
        //cooxml.append(getDataLabelsOOXML(cf));
        // TODO: firstSLiceAng
        cooxml.append("<c:firstSliceAng val=\"0\"/>")
        cooxml.append("<c:holeSize val=\"" + this.getChartOption("donutSize") + "\"/>")

        cooxml.append("</c:doughnutChart>")
        cooxml.append("\r\n")

        return cooxml
    }

}
