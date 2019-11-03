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
import org.json.JSONException
import org.json.JSONObject

import java.util.ArrayList
import java.util.HashMap

open class LineChart(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) : ChartType(charttype, cf, wb) {
    internal var line: Line? = null

    /**
     * return the type JSON for this Chart Object
     *
     * @return
     */
    override/*    	String dojoType;
		if (!this.isStacked()) {
			dojoType="Default";
			typeJSON.put("markers", true);		// default, apparently- doesn't require a MarkerFormat record
		} else
			dojoType="StackedLines";
    	//controls Data Legends, % Distance of sections, Line Format, Area Format, Bar Shapes ...
    	DataFormat df= this.getParentChart().getChartFormat().getDataFormatRec(false);
    	if (df!=null) {
	    	for (int z= 0; z < df.chartArr.size(); z++) {
	            BiffRec b = (BiffRec)df.chartArr.get(z);
	            // LineFormat,AreaFormat,PieFormat,MarkerFormat
	            if (b instanceof Serfmt) {
	            	Serfmt s= (Serfmt) b;
	            	if (s.getSmoothLine()) {
	            		// can be:  Scatter with markers and smoothed lines
	            		typeJSON.put("type", "Default");	// change to Line with Markers
	    				typeJSON.put("markers", true);		// default, apparently- doesn't require a MarkerFormat record
	            	}
	            } else if (b instanceof MarkerFormat) {	// markers for legend (attached label should follow) BUT also can mean NO markers
	            	if (((MarkerFormat)b).getMarkerFormat()==0 && dojoType.equals("Default"))
	            		typeJSON.remove("markers");
	            }
	    	}
    	}
    	typeJSON.put("type", dojoType);*/ val typeJSON: JSONObject
        @Throws(JSONException::class)
        get() = JSONObject()

    init {
        line = charttype as Line
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
            Logger.logErr("Line.getSVG: error in series")
            return ""
        }

        // x value for each point= w/(ncategories + 1) 1st one is xv*2 then increases from there
        // y value for each point= h/YMAX
        val svg = StringBuffer()
        // get data labels, marker formats, series colors
        val n = series.size
        val dls = dataLabelInts // get array of data labels (can be specific per series ...)
        val markers = markerFormats
        var xfactor = 0.0
        var yfactor = 0.0    //
        if (categories!!.size != 0)
            xfactor = w / categories.size    // w/#categories
        else
            xfactor = w
        if (max != 0.0)
            yfactor = h / max    // h/YMAXSCALE
        svg.append("<g>\r\n")
        // define marker shapes for later use
        svg.append(MarkerFormat.markerSVGDefs)
        // for each series
        for (i in 0 until n) {
            // each visible line on the chart consists of two lines- 1 black, 1 color
            var points = ""
            var labels = ""
            val curseries = series[i] as DoubleArray
            val curranges = s.seriesRanges!![i] as Array<String>
            for (j in curseries.indices) {
                points += x + (j + .5) * xfactor + "," + (y + h - curseries[j] * yfactor)
                points += " "
                val l = getSVGDataLabels(dls, axisMetrics, curseries[j], 0.0, i, legends, categories[j].toString())
                if (l != null) {
                    var xx = 2.0 + x + (j + .5) * xfactor
                    if (markers[i] > 0) xx += 10.0    // scoot over for markers
                    labels += "<text x='" + xx + "' y='" + (y + h - curseries[j] * yfactor) +
                            "' " + dataLabelFontSVG + ">" + l + "</text>\r\n"
                }
            }
            // 1st line is black
            svg.append("<polyline " + ChartType.getScript("") + " fill='none' fill-opacity='0' " + ChartType.getStrokeSVG(4f, darkColor) + " points='" + points + "'" + "/>\r\n")
            // 2nd line is the series color
            svg.append("<polyline " + ChartType.getScript("") + "  id='series_" + (i + 1) + "' fill='none' fill-opacity='0' stroke='" + seriescolors!![i] + "' stroke-opacity='1' stroke-width='3' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'" +
                    " points='" + points + "'" +
                    "/>\r\n")
            // Markers, if any, along data points in series
            if (markers[i] > 0) {
                val markerpoints = points.split(" ".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
                for (j in markerpoints.indices) {
                    val markerpoint = markerpoints[j]
                    val xy = markerpoint.split(",".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
                    val xx = java.lang.Double.valueOf(xy[0])
                    val yy = java.lang.Double.valueOf(xy[1])
                    svg.append(MarkerFormat.getMarkerSVG(xx, yy, seriescolors[i], markers[i]) + "\r\n")
                }
            }
            // data labels, if any, after lines and markers
            svg.append(labels)
        }

        svg.append("</g>\r\n")
        return svg.toString()
    }

    /**
     * gets the chart-type specific ooxml representation: <lineChart>
     *
     * @return
    </lineChart> */
    override fun getOOXML(catAxisId: String, valAxisId: String, serAxisId: String): StringBuffer? {
        val cooxml = StringBuffer()

        // chart type: contains chart options and series data
        cooxml.append("<c:lineChart>")
        cooxml.append("\r\n")
        cooxml.append("<c:grouping val=\"")


        if (this.is100PercentStacked)
            cooxml.append("percentStacked")
        else if (this.isStacked)
            cooxml.append("stacked")
        else
            cooxml.append("standard")//			} else if (this.isClustered())
        //				grouping="clustered";
        cooxml.append("\"/>")
        cooxml.append("\r\n")
        // vary colors???

        // *** Series Data:	ser, cat, val for most chart types
        cooxml.append(this.parentChart!!.chartSeries.getOOXML(this.chartType, false, 0))

        // chart data labels, if any
        //TODO: FINISH dlbls
        //cooxml.append(getDataLabelsOOXML(cf));

        //dropLines
        var cl = this.cf!!.getChartLinesRec(ChartLine.TYPE_DROPLINE.toInt())
        if (cl != null)
            cooxml.append(cl.ooxml)
        // hiLowLines
        cl = this.cf!!.getChartLinesRec(ChartLine.TYPE_HILOWLINE.toInt())
        if (cl != null)
            cooxml.append(cl.ooxml)
        // upDownBars
        cooxml.append(cf!!.upDownBarOOXML)
        // marker
        // smooth

        // axis ids	 - unsigned int strings
        cooxml.append("<c:axId val=\"$catAxisId\"/>")
        cooxml.append("\r\n")
        cooxml.append("<c:axId val=\"$valAxisId\"/>")
        cooxml.append("\r\n")

        cooxml.append("</c:lineChart>")
        cooxml.append("\r\n")

        return cooxml
    }
}
