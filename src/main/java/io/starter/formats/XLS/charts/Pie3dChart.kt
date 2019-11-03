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

class Pie3dChart(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) : PieChart(charttype, cf, wb) {

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
            Logger.logErr("Pie.getSVG: error in series")
            return ""
        }
        var dls = dataLabelInts // get array of data labels (can be specific per series ...)
        val LABELOFFSET = 15

        val n = series.size
        var centerx = 0.0
        var centery = 0.0
        var radius = 0.0
        var radiusy = 0.0

        centerx = w / 2 + chartMetrics["x"]
        centery = h / 2 + chartMetrics["y"]
        svg.append("<defs>\r\n")
        svg.append("<filter id=\"multiply\">\r\n")
        svg.append("<feBlend mode=\"multiply\" in2=\"\"/>\r\n")
        svg.append("</filter>\r\n")
        svg.append("</defs>\r\n")
        svg.append("<g>\r\n")

        // if have any series - should :)
        radius = Math.min(w, h) / 2
        //radius= Math.min(w, h);		// ????????????? just an estimate?
        var depth3d = 0.0 // only used for 3d
        radiusy = radius * 2 / 3
        depth3d = radius / 5.5
        if (n > 0) {
            val oneseries = series[0] as DoubleArray    //FOR PIE CHARTS ONLY 1 SERIES VALUE POINTS ARE USED **********************************
            val curranges = s.seriesRanges!![0] as Array<String>
            var total = 0.0
            for (j in oneseries.indices) {    // get total in order to calculate percentages
                total += oneseries[j]
            }
            if (dls.size == 1) {    // no series-specific data labels; expand to entire series for loop below
                val dl = dls[0]
                dls = IntArray(oneseries.size)
                for (i in dls.indices)
                    dls[i] = dl
            }
            var path = ""
            var x = centerx + radius
            var y = centery
            var percentage = 0.0
            var lasta = 0.0
            var largearcflag = 0
            val sweepflag = 0

            // Now create each pie wedge according to it's percentage value
            for (j in oneseries.indices) {
                if (total > 0)
                    percentage = oneseries[j] / total
                val angle = percentage * 360 + lasta
                val x1 = centerx + radius * Math.cos(Math.toRadians(angle))
                val y1 = centery - radiusy * Math.sin(Math.toRadians(angle))
                if (percentage * 360/*angle*/ > 180) {
                    largearcflag = 1
                } else
                    largearcflag = 0

                if (y1 >= centery) {    // the lower portion of a 3d pie chart that makes it look 3d
                    var x0 = x
                    var y0 = y
                    if (centery > y) {    // lower portion has to be below origin
                        y0 = centery
                        x0 = centerx - radius
                    }
                    path = "M" + x0 + "," + y0 + " A" + radius + " " + radiusy + " 0 " + largearcflag + " " + sweepflag + " " + x1 + " " + y1 + " L" + x1 + " " + (y1 + depth3d)
                    path += " A" + radius + " " + radiusy + " 0 " + largearcflag + " " + "1" + " " + x0 + " " + (y0 + depth3d) + " Z"
                    svg.append("<path " + ChartType.getScript(curranges[j]) + "  fill='" + seriescolors!![j] + "' fill-opacity='" + fillOpacity + "' style='filter:url(#multiply)' " + ChartType.getStrokeSVG(2f, mediumColor) + " path='' d='" + path + "' fill-rule='evenodd'/>\r\n")
                }
                path = "M" + centerx + " " + centery + " L" + x + " " + y + " A" + radius + " " + radiusy + " 0 " + largearcflag + " " + sweepflag + " " + x1 + " " + y1 + " L" + centerx + " " + centery + "Z"
                svg.append("<path " + ChartType.getScript(curranges[j]) + "   id='series_" + (j + 1) + "'  fill='" + seriescolors!![j] + "' fill-opacity='1' " + strokeSVG + " path='' d='" + path + "' fill-rule='evenodd'/>\r\n")

                val l = getSVGDataLabels(dls, axisMetrics, oneseries[j], percentage, j, legends, categories!![j].toString())
                if (l != null) {
                    // apparently labels are outside of wedge unless angle is >= 30 ...
                    // category labels
                    val halfa = percentage / 2 * 360 + lasta    // center in area
                    val x2: Double
                    val y2: Double
                    if (percentage < .3) {    // display label on outside with leader lines
                        x2 = centerx + (radius + LABELOFFSET) * Math.cos(Math.toRadians(halfa))
                        y2 = centery - (radiusy + LABELOFFSET) * Math.sin(Math.toRadians(halfa))
                    } else {    // display label within wedge for > 30%
                        x2 = centerx + radius / 2 * Math.cos(Math.toRadians(halfa))
                        y2 = centery - radiusy / 2 * Math.sin(Math.toRadians(halfa))
                    }
                    var style = ""
                    if (percentage >= .3)
                        style = " style='text-anchor: middle;'"
                    else if (lasta > 90 && lasta < 270) {    // right-align text for wedges on left side of pie
                        style = " style='text-anchor: end;'"
                        // TODO: dec x2
                    }
                    svg.append("<text x='$x2' y='$y2' vertical-align='bottom' $dataLabelFontSVG $style>$l</text>\r\n")
                    // leaderline - not exactly like Excel's but ... :) do when NOT putting text within wedge
                    if (percentage < .3) {
                        val x0 = centerx + radius * Math.cos(Math.toRadians(halfa))
                        val y0 = centery - radiusy * Math.sin(Math.toRadians(halfa))
                        svg.append("<line " + ChartType.getScript(curranges[j]) + " x1='" + x0 + "' y1 ='" + y0 + "' x2='" + (x2 - 3) + "' y2='" + (y2 - 3) + "'" + strokeSVG + "/>\r\n")
                    }
                }
                lasta = angle
                x = x1
                y = y1
            }
        }
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
     * gets the chart-type specific ooxml representation: <pie3dChart>
     *
     * @return
    </pie3dChart> */
    override fun getOOXML(catAxisId: String, valAxisId: String, serAxisId: String): StringBuffer? {
        val cooxml = StringBuffer()

        // chart type: contains chart options and series data
        cooxml.append("<c:pie3DChart>")
        cooxml.append("\r\n")
        cooxml.append("<c:varyColors val=\"1\"/>")

        // *** Series Data:	ser, cat, val for most chart types
        cooxml.append(this.parentChart!!.chartSeries.getOOXML(this.chartType, false, 0))

        // chart data labels, if any
        //TODO: FINISH
        //cooxml.append(getDataLabelsOOXML(cf));

        cooxml.append("</c:pie3DChart>")
        cooxml.append("\r\n")

        return cooxml
    }
}


