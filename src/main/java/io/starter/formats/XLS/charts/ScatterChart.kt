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

class ScatterChart(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) : ChartType(charttype, cf, wb) {
    private var scatter: Scatter? = null

    /**
     * return the type JSON for this Chart Object
     *
     * @return
     */
    override/*    	String dojoType;
		if (this.chartType==ChartConstants.SCATTER) {
			dojoType="Default";	// = MarkersOnly
			typeJSON.put("markers", true);
		} else { // Bubble
			dojoType="Bubble";		// "MarkersOnly";
			// shadows: {dx: 2, dy: 2, dw: 2}
		}
    	//controls Data Legends, % Distance of sections, Line Format, Area Format, Bar Shapes ...
    	DataFormat df= this.getParentChart().getChartFormat().getDataFormatRec(false);
    	if (df!=null) {
	    	for (int z= 0; z < df.chartArr.size(); z++) {
	            BiffRec b = (BiffRec)df.chartArr.get(z);
	            if (b instanceof LineFormat) {
	            	if (chartType==ChartConstants.SCATTER)
	                	typeJSON.put("type", "MarkersOnly");	// if Has LineFormat means NO lines!
	            } else if (b instanceof Serfmt) {
	            	Serfmt s= (Serfmt) b;
	            	if (s.getSmoothLine()) {
	            		// can be:  Scatter with markers and smoothed lines
	            		typeJSON.put("type", "Default");	// change to Line with Markers
	    				typeJSON.put("markers", true);		// default, apparently- doesn't require a MarkerFormat record
	            	}
	            }
	        }
    	}
    	typeJSON.put("type", dojoType);
*/ val typeJSON: JSONObject
        @Throws(JSONException::class)
        get() = JSONObject()

    init {
        scatter = charttype as Scatter
    }

    @Throws(JSONException::class)
    fun getJSON(series: Array<ChartSeriesHandle>, wbh: WorkBookHandle, minMax: Array<Double>): JSONObject {
        val chartObjectJSON = JSONObject()

        // Type JSON
        chartObjectJSON.put("type", this.typeJSON)

        // Deal with Series
        var yMax = 0.0
        var yMin = 0.0
        var nSeries = 0
        val seriesJSON = JSONArray()
        val seriesCOLORS = JSONArray()
        var bHasBubbles = false
        try {
            // must trap min and max for axis tick and units
            for (i in series.indices) {
                val seriesvals = CellRange.getValuesAsJSON(series[i].seriesRange, wbh)
                nSeries = Math.max(nSeries, seriesvals.length())
                for (j in 0 until seriesvals.length()) {
                    try {
                        yMax = Math.max(yMax, seriesvals.getDouble(j))
                        yMin = Math.min(yMin, seriesvals.getDouble(j))
                    } catch (n: NumberFormatException) {
                    }

                }
                if (!series[i].hasBubbleSizes())
                    seriesJSON.put(seriesvals)
                else
                    bHasBubbles = true
                seriesCOLORS.put(FormatConstants.SVGCOLORSTRINGS[series[i].seriesColor])
            }
            if (bHasBubbles) {
                // 20080423 KSC: Go thru a second time, after obtaining yMax and yMin, for bubble sizes ...
                for (i in series.indices) {
                    val bubbles = JSONArray()
                    val seriesvals = CellRange.getValuesAsJSON(series[i].seriesRange, wbh)
                    val catvals = CellRange.getValuesAsJSON(series[i].categoryRange, wbh)
                    val bubblesizes = CellRange.getValuesAsJSON(series[i].bubbleSizes, wbh)
                    for (j in 0 until catvals.length()) {
                        val jo = JSONObject()
                        try {
                            jo.put("x", catvals.getDouble(j))
                        } catch (e: Exception) {
                            jo.put("x", j + 1)
                        }

                        jo.put("y", seriesvals.getDouble(j))
                        jo.put("size", Math.round(bubblesizes.getDouble(j) / ((yMax - yMin) / nSeries)))        // TODO: bubble sizes ration is a guess!!
                        bubbles.put(jo)
                    }
                    seriesJSON.put(bubbles)
                }
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
        // x value for each point= w/(ncategories + 1) 1st one is xv*2 then increases from there
        // y value for each point= h/YMAX
        val svg = StringBuffer()

        if (series!!.size == 0) {
            Logger.logErr("Scatter.getSVG: error in series")
            return ""
        }
        // gather data labels, markers, has lines, series colors for chart
        val threeD = cf!!.isThreeD(ChartConstants.SCATTERCHART)
        val dls = dataLabelInts // get array of data labels (can be specific per series ...)
        val n = series.size
        val hasLines = cf!!.hasLines // should be per-series?
        val markers = markerFormats    // get an array of marker formats per series
        if (!hasLines && markers[0] == 0) {
            // if no lines AND no markers set, MUST use default markers (this is what excel does ...)
            val defaultmarkers = intArrayOf(2, 3, 1, 4, 8, 9, 5, 6, 7)
            for (i in 0 until n) {
                markers[i] = defaultmarkers[i]
            }
        }
        /**
         * A Scatter chart has two value axes, showing one set of numerical data along the x-axis
         * and another along the y-axis.
         * It combines these values into single data points and displays them in uneven intervals,
         * or clusters.
         */
        var seriesx: DoubleArray? = null
        var xfactor = 0.0
        var yfactor = 0.0    //
        var TEXTUALXAXIS = true
        // get x axis max/min for an x axis which is a value axis
        var xmin = java.lang.Double.MAX_VALUE
        var xmax = java.lang.Double.MIN_VALUE
        seriesx = DoubleArray(categories!!.size)
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

        if (max != 0.0)
            yfactor = h / max    // h/YMAXSCALE
        svg.append("<g>\r\n")
        // define marker shapes for later use
        svg.append(MarkerFormat.markerSVGDefs)
        // for each series
        for (i in 0 until n) {
            // two lines- 1 black, 1 color
            var points = ""
            var labels = ""
            val seriesy = series[i] as DoubleArray
            for (j in seriesy.indices) {
                var xval = 0.0
                if (TEXTUALXAXIS /*|| i > 0*/)
                    xval = (j + 1).toDouble()
                else
                    xval = seriesx[j]
                points += (x + xval * xfactor).toString() + "," + (y + h - seriesy[j] * yfactor)
                points += " "
                val l = getSVGDataLabels(dls, axisMetrics, seriesy[j], 0.0, i, legends, categories[j].toString())
                if (l != null)
                    labels += "<text x='" + (12.0 + x + xval * xfactor) + "' y='" + (y + h - seriesy[j] * yfactor) +
                            "' " + dataLabelFontSVG + ">" + l + "</text>\r\n"

            }
            if (hasLines) {
                svg.append(getLineSVG(points, seriescolors!![i]))
            }
            // Markers, if any, along data points in series
            if (markers[i] > 0) {
                val markerpoints = points.split(" ".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
                for (j in markerpoints.indices) {
                    val markerpoint = markerpoints[j]
                    val xy = markerpoint.split(",".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
                    val xx = java.lang.Double.valueOf(xy[0])
                    val yy = java.lang.Double.valueOf(xy[1])
                    svg.append(MarkerFormat.getMarkerSVG(xx, yy, seriescolors!![i], markers[i]) + "\r\n")
                }
            }
            // labels after lines and markers
            svg.append(labels)
        }
        svg.append("</g>\r\n")
        return svg.toString()
    }

    /**
     * returns the SVG necessary to define a line at points in color clr
     *
     * @param points String of SVG points
     * @param clr    SVG color String
     * @return
     */
    private fun getLineSVG(points: String, clr: String): String {
        var s = ""
        // each line is comprised of 1 black line and 1 series color line:
        // 1st line is black
        s = "<polyline fill='none' fill-opacity='0' " + ChartType.getStrokeSVG(1f, darkColor) +
                " points='" + points + "'" +
                "/>\r\n"
        // 2nd line is the series color
        s += "<polyline fill='none' fill-opacity='0' " + ChartType.getStrokeSVG(2f, clr) +
                " points='" + points + "'" +
                "/>\r\n"
        return s
    }

    /**
     * gets the chart-type specific ooxml representation: <scatterChart>
     *
     * @return
    </scatterChart> */
    override fun getOOXML(catAxisId: String, valAxisId: String,
                          serAxisId: String): StringBuffer? {
        val cooxml = StringBuffer()

        // chart type: contains chart options and series data
        cooxml.append("<c:scatterChart>")
        cooxml.append("\r\n")
        val markers = markerFormats
        var style: String? = null
        for (m in markers) {
            if (m != 0) {
                if (this.hasSmoothLines)
                    style = "smoothMarker"
                else if (this.hasLines)
                    style = "lineMarker"
                else
                    style = "marker"
                break
            }
        }
        if (style == null && this.hasLines) style = "line"
        if (style == null && this.hasSmoothLines) style = "smooth"
        if (style == null) style = "none"

        cooxml.append("<c:scatterStyle val=\"$style\"/>")
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

        cooxml.append("</c:scatterChart>")
        cooxml.append("\r\n")
        return cooxml
    }
}
