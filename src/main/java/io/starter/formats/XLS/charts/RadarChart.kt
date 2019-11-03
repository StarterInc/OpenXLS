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

open class RadarChart
//private Radar radar = null; can be Radar or RadarArea

(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook)//		radar = (Radar) charttype;
    : ChartType(charttype, cf, wb) {

    var isFilled: Boolean
        get() = chartobj.chartType == ChartConstants.RADARAREACHART
        set(isFilled) {
            if (isFilled)
                chartobj = RadarArea.prototype as RadarArea?
            else
                chartobj = Radar.prototype as Radar?
            chartobj.parentChart = cf!!.parentChart
            cf!!.chartArr.removeAt(0)
            cf!!.chartArr.add(chartobj)
        }


    /**
     * returns SVG to represent the actual chart object i.e. the representation
     * of the series data in the particular format (BAR, LINE, AREA, etc.)
     *
     * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min,
     * max
     * @param axisMetrics  maps specific axis options such as xAxisReversed, xPattern ...
     * @param categories
     * @param series       arraylist of double[] series values
     * @param seriescolors int[] of series or bar colors color ints
     * @return String svg
     */
    override fun getSVG(chartMetrics: HashMap<String, Double>,
                        axisMetrics: HashMap<String, Any>, s: ChartSeries): String {
        val x = chartMetrics["x"] // + (!yAxisReversed?0:w); // x is
        // constant at x origin unless
        // reversed
        val y = chartMetrics["y"]
        val w = chartMetrics["w"]
        val h = chartMetrics["h"]
        val max = chartMetrics["max"]
        val min = chartMetrics["min"]
        val categories = s.getCategories()
        val series = s.seriesValues
        val seriescolors = s.seriesBarColors
        val legends = s.getLegends()
        // x value for each point= w/(ncategories + 1) 1st one is xv*2 then
        // increases from there
        // y value for each point= h/YMAX
        if (series!!.size == 0) {
            Logger.logErr("Radar.getSVG: error in series")
            return ""
        }
        val svg = StringBuffer()

        // obtain data label, marker info for chart + series colors
        val dls = dataLabelInts // get array of data labels (can be
        // specific per series ...)
        val markers = markerFormats // get an array of marker formats
        // per series
        val n = series.size
        val nseries = categories!!.size

        val centerx = w / 2 + x
        val centery = h / 2 + y
        val percentage = 1.0 / nseries // divide into equal sections
        val radius = Math.min(w, h) / 2.3 // should take up almost entire
        // w/h of chart

        svg.append("<g>\r\n")
        // define marker shapes for later use
        svg.append(MarkerFormat.markerSVGDefs)
        // for each series
        for (i in 0 until n) {
            var points = ""
            var labels = ""
            var angle = 90.0 // starts straight up
            val curseries = series[i] as DoubleArray
            val curranges = s.seriesRanges!![i] as Array<String>
            var x0 = 0.0
            var y0 = 0.0
            for (j in curseries.indices) {
                // get next point as a percentage of the radius
                val r = radius * (curseries[j] / (max - min))
                val x1 = centerx + r * Math.cos(Math.toRadians(angle)) //
                val y1 = centery - r * Math.sin(Math.toRadians(angle))
                if (j == 0) { // save initial points so can close the loop at
                    // end
                    x0 = x1
                    y0 = y1
                }
                points += "$x1,$y1 "
                val l = getSVGDataLabels(dls, axisMetrics, curseries[j],
                        percentage, j, legends, categories[j].toString())
                if (l != null) {
                    val labelx1 = centerx + (r + 5) * Math.cos(Math.toRadians(angle)) //
                    val labely1 = centery - (r + 5) * Math.sin(Math.toRadians(angle))
                    labels += ("<text x='" + labelx1 + "' y='" + labely1
                            + "' style='text-anchor: middle;' "
                            + dataLabelFontSVG + ">" + l + "</text>\r\n")
                }
                angle -= percentage * 360 // next point on next category
                // radial line
            }
            // close loop
            points += "$x0,$y0"
            // 1st line is black
            svg.append("<polyline " + ChartType.getScript("") + " fill-opacity='0' "
                    + ChartType.getStrokeSVG(4f, "black") + " points='" + points + "'"
                    + "/>\r\n")
            // 2nd line is the series color
            svg.append("<polyline " + ChartType.getScript("") + "   id='series_"
                    + (i + 1) + "' fill='none' fill-opacity='0' "
                    + ChartType.getStrokeSVG(3f, seriescolors!![i]) + " points='" + points
                    + "'" + "/>\r\n")
            // Markers, if any, along data points in series
            if (markers[i] > 0) {
                val markerpoints = points.split(" ".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
                for (j in markerpoints.indices) {
                    val markerpoint = markerpoints[j]
                    val xy = markerpoint.split(",".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
                    val xx = java.lang.Double.valueOf(xy[0])
                    val yy = java.lang.Double.valueOf(xy[1])
                    svg.append(MarkerFormat.getMarkerSVG(xx, yy,
                            seriescolors[i], markers[i]) + "\r\n")
                }
            }
            // labels after lines and markers
            svg.append(labels)
        }

        svg.append("</g>\r\n")
        return svg.toString()
    }

    /**
     * gets the chart-type specific ooxml representation: <radarChart>
     *
     * @return
    </radarChart> */
    override fun getOOXML(catAxisId: String, valAxisId: String, serAxisId: String): StringBuffer? {
        val cooxml = StringBuffer()

        // chart type: contains chart options and series data
        cooxml.append("<c:radarChart>")
        cooxml.append("\r\n")

        var style = "standard"
        if (this.isFilled)
            style = "filled"
        else {
            val markers = markerFormats
            for (m in markers) {
                if (m != 0) {
                    style = "marker"
                    break
                }
            }
        }
        cooxml.append("<c:radarStyle val=\"$style\"/>")
        // vary colors???

        // *** Series Data: ser, cat, val for most chart types
        cooxml.append(this.parentChart!!.chartSeries.getOOXML(this.chartType, false, 0))

        // chart data labels, if any
        // TODO: FINISH
        // cooxml.append(getDataLabelsOOXML(cf));

        // axis ids - unsigned int strings
        cooxml.append("<c:axId val=\"$catAxisId\"/>")
        cooxml.append("\r\n")
        cooxml.append("<c:axId val=\"$valAxisId\"/>")
        cooxml.append("\r\n")

        cooxml.append("</c:radarChart>")
        cooxml.append("\r\n")
        return cooxml
    }
}
