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

class OfPieChart(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) : ChartType(charttype, cf, wb) {
    internal var ofPie: Boppop? = null

    init {
        ofPie = charttype as Boppop
    }

    /**
     * gets the chart-type specific ooxml representation: <ofPieChart>
     *
     * @return
    </ofPieChart> */
    override fun getOOXML(catAxisId: String, valAxisId: String, serAxisId: String): StringBuffer? {
        val cooxml = StringBuffer()

        // chart type: contains chart options and series data
        cooxml.append("<c:ofPieChart>")
        cooxml.append("\r\n")
        cooxml.append("<c:ofPieType val=\"" + (if (ofPie!!.isPieOfPie) "pie" else "bar") + "\"/>")
        cooxml.append("<c:varyColors val=\"1\"/>")

        // *** Series Data:	ser, cat, val for most chart types
        cooxml.append(this.parentChart!!.chartSeries.getOOXML(this.chartType, false, 0))

        // chart data labels, if any
        //TODO: FINISH
        //cooxml.append(getDataLabelsOOXML(cf));
        // gapWidth
        if (ofPie!!.getpcGap() != 150)
            cooxml.append("<c:gapWidth val=\"" + ofPie!!.getpcGap() + "\"/>")
        // splitType
        if (ofPie!!.splitType != -1)
            cooxml.append("<c:splitType val=\"" + ofPie!!.splitTypeOOXML + "\"/>")
        // splitPos
        if (ofPie!!.splitPos != 0)
            cooxml.append("<c:splitPos val=\"" + ofPie!!.splitPos + "\"/>")
        // custSplit TODO: FINISH
        // secondPieSize
        if (ofPie!!.secondPieSize != 75)
            cooxml.append("<c:secondPieSize val=\"" + ofPie!!.secondPieSize + "\"/>")
        // serLines
        val cl = cf!!.chartLinesRec
        if (cl != null)
            cooxml.append(cl.ooxml)
        cooxml.append("</c:ofPieChart>")
        cooxml.append("\r\n")

        return cooxml
    }
}
