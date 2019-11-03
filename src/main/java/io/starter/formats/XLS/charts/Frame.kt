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

import io.starter.formats.XLS.BiffRec
import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools

/**
 * **Frame: Defines Border Shape Around Displayed Text (0x1032)**
 */

/**
 * frt (2 bytes): An unsigned integer that specifies the type of frame to be drawn. MUST be a value from the following table:
 * Type of frame
 * 0x0000	    A frame surrounding the chart element.
 *
 * 0x0004		A frame with shadow surrounding the chart element.
 *
 * A - fAutoSize (1 bit): A bit that specifies if the size of the frame is automatically calculated. If the value is 1, the size of the frame is automatically calculated. In this case, the width and height specified by the chart element are ignored and the size of the frame is calculated automatically. If the value is 0, the width and height specified by the chart element are used as the size of the frame.
 *
 * B - fAutoPosition (1 bit): A bit that specifies if the position of the frame is automatically calculated. If the value is 1, the position of the frame is automatically calculated. In this case, the (x, y) specified by the chart element are ignored, and the position of the frame is automatically calculated. If the value is 0, the (x, y) location specified by the chart element are used as the position of the frame.
 */
class Frame : GenericChartObject(), ChartObject {
    internal var fAutoSize: Boolean = false
    internal var fAutoPosition: Boolean = false
    internal var frt: Int = 0

    /**
     * return the bg color assoc with this frame rec
     * in main chart record
     * NOTE that bg color is defined by the frame rec's associated
     * AreaFormat Record's FOREGROUND color (icvFore)
     * @return bg color Hex String
     */
    val bgColor: String?
        get() = getBgColor(chartArr)

    /**
     * return the line color as a hex String
     * @return
     */
    val lineColor: String?
        get() {
            val l = Chart.findRec(this.chartArr, LineFormat::class.java) as LineFormat
            return l?.lineColor
        }

    override fun init() {
        frt = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        val flag = this.getByteAt(2)// 3= autosize, autoposition
        fAutoSize = flag and 0x01 == 0x01
        fAutoPosition = flag and 0x02 == 0x02
        super.init()
    }

    /**
     * sets the background color assoc with this frame rec
     * NOTE that bg color is defined by the frame rec's associated
     * AreaFormat Record's FOREGROUND color (icvFore)
     * @param bg  color int
     */
    fun setBgColor(bg: Int) {
        for (i in chartArr.indices) {
            val b = chartArr[i]
            if (b is AreaFormat)
                b.seticvFore(bg)
        }
    }

    /**
     * set the frame autosize/autoposition, necessary for legend expansion
     * @see Legend.setAutoPosition, Pos.setAutosizeLegend
     * [BugTracker 2844]
     */
    fun setAutosize() {
        this.data[2] = 3 // sets both fAutoSize and fAutoPostion
    }

    /**
     * adds a frame box with the desired lineweight and fill color
     * @param lw    if -1, none
     * @param lclr
     * @param bgcolor    if -1, none
     */
    fun addBox(lw: Int, lclr: Int, bgcolor: Int) {
        var lf = Chart.findRec(chartArr, LineFormat::class.java) as LineFormat
        if (lf == null) {
            lf = LineFormat.getPrototype(0, 0) as LineFormat
            this.addChartRecord(lf)
        }
        if (lw != -1) {
            lf.setLineWeight(lw)
            lf.lineStyle = 0    // solid
        } else {
            lf.lineStyle = 5    // none
        }
        if (lclr != -1)
            lf.setLineColor(lclr)
        var af = Chart.findRec(chartArr, AreaFormat::class.java) as AreaFormat
        if (af == null) {
            af = AreaFormat.getPrototype(1) as AreaFormat
            this.addChartRecord(af)
        }
        if (bgcolor != -1)
            af.seticvBack(bgcolor)
    }

    /**
     * returns true if this Frame is surrounded by a box
     * @return
     */
    fun hasBox(): Boolean {
        val l = Chart.findRec(this.chartArr, LineFormat::class.java) as LineFormat
        return l.lineStyle != LineFormat.NONE
    }

    /**
     * return the svg representation of this Frame element
     * @param coords
     * @return
     */
    fun getSVG(coords: FloatArray): StringBuffer {
        val svg = StringBuffer()
        var lineSVG = ""
        var bgclr = bgColor
        //		String bgclr= "white";
        if (bgclr == null)
            bgclr = "white"

        val lf = Chart.findRec(chartArr, LineFormat::class.java) as LineFormat
        if (lf != null)
            lineSVG = lf.svg

        val x = coords[0] - coords[2] / 2    // apparently coords are center-point; adjust
        val y = coords[1] - coords[3] / 2
        svg.append("<rect x='" + x + "' y='" + y + "' width='" + coords[2] + "' height='" + coords[3] +
                "' fill='" + bgclr + "' fill-opacity='1' " + lineSVG + "/>\r\n")

        return svg
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 6302152932127918650L

        val prototype: XLSRecord?
            get() {
                val f = Frame()
                f.opcode = XLSConstants.FRAME
                f.data = byteArrayOf(0, 0, 2, 0)
                f.init()
                return f
            }


        /**
         * static utility to return the bg color assoc with the desired object
         * <br></br>first checks if a gelFrame record exists; if so, it uses that color.
         * <br></br>if no gelFrame record exists, it looks for an AreaFormat record.
         * @return bg color Hex String
         */
        fun getBgColor(chartArr: java.util.ArrayList<*>): String? {
            val gf = Chart.findRec(chartArr, GelFrame::class.java) as GelFrame
            if (gf != null)
                return gf.fillColor
            val af = Chart.findRec(chartArr, AreaFormat::class.java) as AreaFormat
            return af?.fillColorStr
// use default
        }
    }
}
