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

//OOXML-specific structures

import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.OOXML.Layout
import io.starter.formats.OOXML.SpPr
import io.starter.formats.OOXML.Title
import io.starter.formats.OOXML.TxPr
import io.starter.toolkit.Logger

import java.util.ArrayList
import java.util.HashMap


class OOXMLChart
/**
 * create a new OOXMLChart from a 2003-v chart object
 *
 * @param c
 * @param wbh
 */
(c: Chart, wbh: WorkBookHandle) : Chart() {
    var lang = "en-US"                                // default
    var roundedCorners = false
    /**
     * return the OOXML title element for this chart
     *
     * @return
     */
    var ooxmlTitle: Title? = null                                    // title element
    var ooxmlLegend: io.starter.formats.OOXML.Legend? = null
    var plotAreaLayout: Layout? = null
    private var plotareashapeProps: SpPr? = null                    // defines the shape properties for this chart (line and fill)
    private var csshapeProps: SpPr? = null                        // defines the shape properties for this chart space
    /**
     * return the OOXML text formatting element for this chart, if present
     *
     * @return
     */
    /**
     * store the OOXML text formatting element for this chart
     */
    var txPr: TxPr? = null
    /**
     * return state of  how to resize or move upon edit OOXML specific
     *
     * @return editMovement string
     */
    /**
     * specify how to resize or move upon edit OOXML specific
     *
     * @param editMovement
     */
    var editMovement = "twoCell"
        set(editMovement) {
            field = editMovement
            dirtyflag = true
        }                    // default
    private var name: String? = null                                // name property of cNvPr
    /**
     * returns information regarding external files associated with this chart
     * <br></br>e.g. a chart user shape, an image
     *
     * @return
     */
    var chartEmbeds: ArrayList<*>? = null
        private set                    // if present, name(s) of chart embeds (images, userShape definitions (drawingml files that define shapes ontop of chart))

    /**
     *
     */
    /**
     * set the OOXML-specific name for this chart
     *
     * @param name
     */
    var ooxmlName: String?
        get() = this.name
        set(name) {
            this.name = name
            dirtyflag = true
        }

    init {
        // Walk up the superclass hierarchy
        //io.starter.toolkit.Logger.log("BEFORE: chartArr: " + Arrays.toString(chartArr.toArray()));
        run {
            var obj: Class<*> = c.javaClass
            while (obj != Any::class.java) {
                val fields = obj.declaredFields
                for (i in fields.indices) {
                    fields[i].isAccessible = true
                    try {
                        // for each class/suerclass, copy all fields
                        // from this object to the clone
                        fields[i].set(this, fields[i].get(c))
                    } catch (e: IllegalArgumentException) {
                    } catch (e: IllegalAccessException) {
                    }

                }
                obj = obj.superclass
            }
        }


        // ttl?
        // chartlaout?
        // txpr?
        // name?
        this.name = c.title
        if (c.hasDataLegend()) {
            ooxmlLegend = io.starter.formats.OOXML.Legend.createLegend(c.legend)
        }
        this.wbh = wbh
    }

    override fun toString(): String {
        val t = title
        return if (t != "") t else name
// if no title, return OOXMLName
    }


    /**
     * return the OOXML shape property for this chart
     *
     * @return
     * @param    type    0= chart shape props, 1=plot area shape props 2= chartspace shape props, 3= legend shape props
     */
    fun getSpPr(type: Int): SpPr? {
        if (type == 0)
            return plotareashapeProps
        return if (type == 1) csshapeProps else null
    }

    /**
     * define the OOXML shape property for this chart from an existing spPr element
     *
     * @param type 0=plot area shape props 1= chartspace shape props, 2= legend shape props
     */
    fun setSpPr(type: Int, spPr: SpPr) {
        if (type == 0) {
            plotareashapeProps = spPr    // plot area
            var lw = -1
            var lclr = 0
            val bgcolor = -1
            lw = spPr.lineWidth    // TO DO: Style
            lclr = spPr.lineColor
            //bgcolor= spPr.getColor();
            this.axes!!.setPlotAreaBorder(lw, lclr)
        } else if (type == 1) {
            csshapeProps = spPr
        }

    }

    /**
     * set the OOXML title element for this chart
     *
     * @param t
     */
    fun setOOXMLTitle(t: Title, wb: WorkBookHandle) {
        ooxmlTitle = t
        var fid = ooxmlTitle!!.getFontId(wb)
        if (fid == -1) fid = 5    // default ...?
        var coords: FloatArray? = null
        var lw = -1
        var lclr = 0
        var bgcolor = 0
        if (ooxmlTitle!!.layout != null) {    // pos
            coords = ooxmlTitle!!.layout!!.coords
        }
        if (ooxmlTitle!!.spPr != null) {    // Area Fill, Line Format
            val sp = ooxmlTitle!!.spPr
            lw = sp!!.lineWidth    // TO DO: Style, fill/color ...
            lclr = sp.lineColor
            bgcolor = sp.color
        }
        if (coords != null)
            titleTd!!.setFrame(lw, lclr, bgcolor, coords)

        // must also set the fontx id for the title
        if (titleTd != null)
            titleTd!!.fontId = fid

    }

    /**
     * remove the legend from the chart
     */
    override fun removeLegend() {
        showLegend(false, false)
        ooxmlLegend = null
    }

    /**
     * sets external information linked to or "embedded" in this OOXML chart;
     * can be a chart user shape, an image ...
     * <br></br>NOTE: a userShape is a drawingml file name which defines the userShape (if any)
     * <br></br>a userShape is a drawing or shape ontop of a chart
     *
     * @param String[] embedType, filename e.g. {"userShape", "userShape file name"}
     */
    fun addChartEmbed(ce: Array<String>) {
        if (chartEmbeds == null)
            chartEmbeds = ArrayList()
        chartEmbeds!!.add(ce)
    }

    /**
     * return the OOXML representation of this chart object "c:chart" representing OOXML element in chartX.xml
     * <br></br>below is complete the ordered sequence of child elements of the chart element
     * <br></br>c:chart - parent= chartSpace
     *  * title
     *  * autoTitleDeleted
     *  * pivotFmts
     *  * view3d
     *  * floor
     *  * sideWall
     *  * backWall
     *  * plotArea (see below)
     *  * legend
     *  * plotVisOnly
     *  * dispBlankAs
     *  * showDlblsOverMax
     *
     * <br></br>plotArea:
     *  * layout
     *  * chart type (see below)
     *  *  axes ***
     *  * dTable
     *  * spPr
     *
     * <br></br>chart type:
     *  * barDir		Bar, Bar3d only
     *  * radarStyle || scatterStyle
     *  * ofPieType
     *  * wireFrame	surface
     *  * grouping	Area, Area3d, Line, Line3d, Bar, Bar3D
     *  * varyColors		not for Stock
     *  * ser  *n series
     *  * dLbls			not for surface
     * Area Chart, AreaChart3D, LineChart, Line3D, Stock
     *  * dropLines
     * Bar Chart, Bar3d, ofPieChart
     *  * gapWidth
     * AreaChart3D, Line3D, Bar3D
     *  * gapDepth
     * Line, Stock
     *  * hiLowLines
     *  * upDownBars
     * Line
     *  * marker
     *  * smooth
     * BarChart only
     *  * overlap
     *  * serLines
     * Bar 3d only
     *  * shape
     * ofPieChart
     *  * splitType
     *  * splitPos
     *  * custSplit
     *  * secondPieSize
     *  * serLines
     * Pie, Doughnut
     *  * firstSliceAng
     * Doughnut
     *  * holeSize
     * Surface
     *  * bandFmts
     * Bubble
     *  * bubble3D
     *  * bubbleScale
     *  * showNegBubbles
     *  * sizeRepresents
     *  * axId
     *
     * @return StringBuffer
     */
    fun getOOXML(catAxisId: String, valAxisId: String, serAxisId: String): StringBuffer {
        val cooxml = StringBuffer()
        try {
            val allCharts = this.allChartTypes        // usually only 1 chart but some have overlay charts in addition to the default chart (chart 0)

            // lang
            cooxml.append("<c:lang val=\"" + this.lang + "\"/>")
            cooxml.append("\r\n")
            // rounded corners
            if (this.roundedCorners)
                cooxml.append("<c:roundedCorners val=\"1\"/>")
            // chart
            cooxml.append("<c:chart>")
            cooxml.append("\r\n")
            // title
            if (this.ooxmlTitle == null) {// if no OOXML title, see if have a BIFF8 title
                if (this.title != "") {
                    this.setOOXMLTitle(Title(this.titleTd!!, this.wbh!!.workBook), this.wbh)
                }
            }
            if (this.ooxmlTitle != null)
            // otherwise there's no title
                cooxml.append(this.ooxmlTitle!!.ooxml)

            if (allCharts[0] != ChartConstants.BUBBLECHART) {    // bubble threeD handled in series for some reason
                // Q: what if overlay charts are not 3D?  what if default isn't and overlay is?
                // view 3D
                val td = this.getThreeDRec(0)
                if (td != null)
                    cooxml.append(td.ooxml)
            }
            // TODO: Handle:
            // floor
            // sideWall
            // backWall
            // plot area (contains all chart types such as barChart, lineChart ...)
            cooxml.append("<c:plotArea>")
            cooxml.append("\r\n")
            // layout: size and position
            if (this.plotAreaLayout == null) {    // if converted from XLS will hit here
                val chartMetrics = this.getMetrics(wbh)
                val x = chartMetrics["x"] / chartMetrics["canvasw"]
                val y = chartMetrics["y"] / chartMetrics["canvash"]
                val w = chartMetrics["w"] / chartMetrics["canvasw"]
                val h = chartMetrics["h"] / chartMetrics["canvash"]
                this.plotAreaLayout = Layout("inner", doubleArrayOf(x, y, w, h))
            }
            cooxml.append(this.plotAreaLayout!!.ooxml)

            for (ch in chartgroup) {
                cooxml.append(ch.getOOXML(catAxisId, valAxisId, serAxisId))
            }

            /* TODO:
             * 1- varyColors  ???
             * 2- getChartSeries.getOOXML  -> nchart?
             * 3- data labels
             * 4- drop lines
             *
             * area charts -- bar colors!!!
             * bar - serLines
             * pie, doughnut - firstSliceArg
             * radar -- filled?
             * bubble	-- bubbleScale, showNegBubbles, sizeRepresents
             * surface -- wireframe, bandfmts
             *
             * pie of pie, bar of pie
             * surface3d
             * stock
             *
             *
             */

            // ******************************************************************************
            // after chart type ooxml, axes (if present)
            cooxml.append(this.axes!!.getOOXML(ChartConstants.XAXIS, 0, catAxisId, valAxisId))
            cooxml.append(this.axes!!.getOOXML(ChartConstants.XVALAXIS, 2, catAxisId, valAxisId))    // valAx - for bubble/scatter
            // val axis
            cooxml.append(this.axes!!.getOOXML(ChartConstants.YAXIS, 1, valAxisId, catAxisId))    // val axis
            // ser axis
            cooxml.append(this.axes!!.getOOXML(ChartConstants.ZAXIS, 3, serAxisId, valAxisId))    // ser axis (crosses val axis)
            // TODO: dateAx
            if (this.getSpPr(0) != null) {    // plot area shape props
                cooxml.append(this.getSpPr(0)!!.ooxml)
            } else if (!this.wbh!!.isExcel2007) {
                val sp = SpPr("c", this.plotAreaBgColor.substring(1), 12700, this.plotAreaLnColor.substring(1))
                cooxml.append(sp.ooxml)

            }

            cooxml.append("</c:plotArea>")
            cooxml.append("\r\n")
            // legend
            if (this.ooxmlLegend != null) cooxml.append(this.ooxmlLegend!!.ooxml)

            cooxml.append("<c:plotVisOnly val=\"1\"/>")        // specifies that only visible cells should be plotted on the chart
            //	    	<c:dispBlanksAs val="gap"/>	"gap", "span", "zero"  --> default
            cooxml.append("\r\n")
            cooxml.append("</c:chart>")
            cooxml.append("\r\n")
            if (this.getSpPr(1) != null) { // chart space shape props
                cooxml.append(this.getSpPr(1)!!.ooxml)
            }
            if (this.txPr != null) { // text formatting
                cooxml.append(this.txPr!!.ooxml)
            }
        } catch (e: Exception) {
            Logger.logErr("OOXMLChart.getOOXML: error generating OOXML.  Chart not created: $e")
        }

        return cooxml    //.toString();
    }

}
