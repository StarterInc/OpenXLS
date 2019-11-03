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

class Area3DChart(charttype: GenericChartObject, cf: ChartFormat, wb: WorkBook) : AreaChart(charttype, cf, wb) {
    init {
        area = charttype as Area
    }

    /**
     * gets the chart-type specific ooxml representation: <areaChart>
     *
     * @return
    </areaChart> */
    override fun getOOXML(catAxisId: String, valAxisId: String, serAxisId: String): StringBuffer? {
        val cooxml = StringBuffer()

        // chart type: contains chart options and series data
        cooxml.append("<c:area3DChart>")
        cooxml.append("\r\n")
        cooxml.append("<c:grouping val=\"")

        if (this.is100PercentStacked)
            cooxml.append("percentStacked")
        else if (this.isStacked)
            cooxml.append("stacked")
        else
            cooxml.append("standard")
        cooxml.append("\"/>")
        cooxml.append("\r\n")
        // vary colors???

        // *** Series Data:	ser, cat, val for most chart types
        cooxml.append(this.parentChart!!.chartSeries.getOOXML(this.chartType, false, 0))

        // chart data labels, if any
        //TODO: FINISH
        //cooxml.append(getDataLabelsOOXML(cf));
        // TODO: get real value
        val gapdepth = this.gapDepth // 150 is default
        if (gapdepth != 0 && gapdepth != 150)
            cooxml.append("<c:gapDepth val=\"$gapdepth\"/>")
        //DropLines
        val cl = cf!!.chartLinesRec
        if (cl != null)
            cooxml.append(cl.ooxml)

        // axis ids	 - unsigned int strings
        cooxml.append("<c:axId val=\"$catAxisId\"/>")
        cooxml.append("\r\n")
        cooxml.append("<c:axId val=\"$valAxisId\"/>")
        cooxml.append("\r\n")
        cooxml.append("<c:axId val=\"$serAxisId\"/>")
        cooxml.append("\r\n")

        cooxml.append("</c:area3DChart>")
        cooxml.append("\r\n")

        return cooxml
    }

}