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

import io.starter.OpenXLS.FormatHandle
import io.starter.formats.OOXML.SpPr
import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools

/**
 * **LineFormat: Style of a Line or border(0x1007)**
 *
 *
 * 4		rgb		4		Color of line: high byte must be 0
 * 8		lnx		2		Pattern of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
 * When the value of this field is 0x0005 (None), the values of we and icv MUST be set to: Line thickness (we)= 0xFFFF (Hairline)   Line color (icv)  0x004D
 * 10		we		2		Weight of line	(-1= hairline, 0= narrow, 1= med (double), 2= wide (triple)
 * 12		grbit	2		flags
 * 14		icv				Index to color of line.
 * An Icv that specifies a color from the chart color table. This value MUST be greater than or equal to 0x0000 and less than or equal to 0x0041,
 * or greater than or equal to 0x004D and less than or equal to 0x00004F. This value SHOULD NOT be less than 0x0008.
 * 0x0040 == Default foreground color. This is the window text color in the sheet display.
 * 0x0041 == Default background color. This is the window background color in the sheet display and is the default background color for a cell.
 * 0x004D == Default chart foreground color. This is the window text color in the chart display.
 *
 *
 * grbit:
 * 0		0x1		fAuto		Automatic format
 * 1				reserved, 0
 * 2		0x4		fAxisOn		specifies whether axis line is displayed
 * If the previous record is AxisLine and the value of the id field of the AxisLine record is equal to 0x0000, this field MUST be a value from the following table:
 * fAxisOn			Lns								Meaning
 * 0				0x0005							The axis line is not displayed.
 * 0				Any legal value except 0x0005	The axis line is displayed.
 * 1				Any legal value					The axis line is displayed.
 * If the previous record is not AxisLine and the value of the id field of the AxisLine record is equal to 0x0000, this field MUST be zero, and MUST be ignored.
 * 3		0x8		fAutoColor		specifies whether icv= 0x4D.  if 1, icv must= 0x4D.
 */
class LineFormat : GenericChartObject(), ChartObject {
    private var rgb: java.awt.Color? = null
    private var lnx: Short = 0
    private var we: Short = 0
    private var grbit: Short = 0
    private var icv: Short = 0
    private var sppr: SpPr? = null

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 0, 0, 0, 0, -1, -1, 9, 0, 77, 0)    // no line default
    // TODO: Figure this out!!
    private val PROTOTYPE_BYTES_1 = byteArrayOf(-128, -128, -128, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    /**
     * return Pattern of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
     */
    /**
     * Style of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
     *
     * @param style
     */
    //  When the value of this field is 0x0005 (None), the values of we and icv MUST be set to: Line thickness (we)= 0xFFFF (Hairline)   Line color (icv)  0x004D
    //* 10		we		2		Weight of line	(-1= hairline, 0= narrow, 1= med (double), 2= wide (triple)
    // auto color
    var lineStyle: Int
        get() = lnx.toInt()
        set(style) {
            lnx = style.toShort()
            if (lnx.toInt() == 5) {
                we = -1
                grbit = 0x8
                setLineColor(0x4D)
            }
            updateRecord()
        }

    /**
     * return the line color as a hex String
     *
     * @return
     */
    val lineColor: String
        get() = FormatHandle.colorToHexString(rgb!!)

    // no line
    // hairline
    // narrow	- rest are just guesses really
    // medium
    // wide
    // dark grey pattern
    // medium grey pattern
    // light grey pattern
    val svg: String
        get() {
            if (lnx.toInt() == 5)
                return ""
            var sz = 1f
            if (we.toInt() == -1)
                sz = 1f
            else if (we.toInt() == 0)
                sz = 2f
            else if (we.toInt() == 1)
                sz = 4f
            else if (we.toInt() == 2)
                sz = 6f
            var clr = ChartType.mediumColor
            if (lnx.toInt() == DKGRAY)
                clr = ChartType.darkColor
            else if (lnx.toInt() == MEDGRAY)
                clr = ChartType.mediumColor
            else if (lnx.toInt() == LTGRAY)
                clr = ChartType.lightColor
            var style = ""
            if (lnx.toInt() == DASH) {
                style = " style='stroke-dasharray: 9, 5;' "
            } else if (lnx.toInt() == DOT) {
                style = " style='stroke-dasharray:2, 2;' "
            } else if (lnx.toInt() == DASHDOT) {
                style = " style='stroke-dasharray: 3, 2, 9, 2;' "
            } else if (lnx.toInt() == DASHDASHDOT) {
                style = " style='stroke-dasharray: 9, 5, 9, 5, 3, 2;' "
            }
            return " stroke='" + clr + "'  stroke-opacity='1' stroke-width='" + sz + "' " + style + "stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'"
        }

    /**
     * if have 2007+v settings, use them.  Otherwise, interpret line settings to OOXML
     *
     * @return
     */
    // TODO:  if changed lineformat info, change in sp
    // TODO: line styles + weight
    val ooxml: String
        get() {
            if (sppr != null) {
                return sppr!!.ooxml
            }

            if (!parentChart!!.workBook!!.isExcel2007) {
                val ooxml = StringBuffer()
                ooxml.append("<c:spPr>")
                ooxml.append("<a:ln w=\"$we\">")
                ooxml.append("<a:solidFill>")
                ooxml.append("<a:srgbClr val=\"" + FormatHandle.colorToHexString(rgb!!) + "\"/>")
                ooxml.append("</a:solidFill>")
                ooxml.append("</a:ln>")
                ooxml.append("</c:spPr>")
                return ooxml.toString()
            }
            return ""
        }

    override fun init() {
        super.init()
        val data = this.data
        rgb = java.awt.Color(if (data!![0] < 0) 255 + data[0] else data[0], if (data[1] < 0) 255 + data[1] else data[1], if (data[2] < 0) 255 + data[2] else data[2])
        lnx = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        we = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        grbit = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())
        icv = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt())
    }

    /**
     * 10		we		2		Weight of line	(-1= hairline, 0= narrow, 1= med (double), 2= wide (triple)
     */
    fun setLineWeight(weight: Int) {
        we = weight.toShort()
        updateRecord()
    }


    /**
     * Index to color of line
     *
     * @param clr
     */
    fun setLineColor(clr: Int) {
        if (clr > -1 && clr < this.colorTable.size) {
            icv = clr.toShort()
            rgb = this.colorTable[clr]
            updateRecord()
        } else if (clr == 0x4D) { // special flag, default fg
            icv = clr.toShort()
            rgb = this.colorTable[0]
            updateRecord()
        }
        // TOOD: finish
        if (sppr != null) {
            //    		sppr.setLine(we, clr);
        }
    }

    /**
     * update the underlying data
     */
    private fun updateRecord() {
        var b = ByteArray(4)
        b[0] = rgb!!.red.toByte()
        b[1] = rgb!!.green.toByte()
        b[2] = rgb!!.blue.toByte()
        b[3] = 0    // reserved/0
        System.arraycopy(b, 0, this.data!!, 0, 4)
        b = ByteTools.shortToLEBytes(lnx)
        this.data[4] = b[0]
        this.data[5] = b[1]
        b = ByteTools.shortToLEBytes(we)
        this.data[6] = b[0]
        this.data[7] = b[1]
        b = ByteTools.shortToLEBytes(grbit)
        this.data[8] = b[0]
        this.data[9] = b[1]
        b = ByteTools.shortToLEBytes(icv)
        this.data[10] = b[0]
        this.data[11] = b[1]
    }

    override fun toString(): String {
        return "LineFormat: LinePattern=" + lnx + " Weight=" + we + " Draw Ticks=" + (grbit and 0x4 == 0x4)
    }

    /**
     * sets the OOXML settings for this line
     *
     * @param sp
     */
    fun setFromOOXML(sp: SpPr) {
        this.sppr = sp
        val lw = sp.lineWidth
        this.setLineWeight(lw)        // sp lw in emus.  1 pt= 12700 emus.
        this.setLineColor(sp.lineColor)
        this.lineStyle = sp.lineStyle
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 3051781109844837056L
        val SOLID = 0
        val DASH = 1
        val DOT = 2
        val DASHDOT = 3
        val DASHDASHDOT = 4
        val NONE = 5
        val DKGRAY = 6
        val MEDGRAY = 7
        val LTGRAY = 8

        val prototype: XLSRecord?
            get() {
                val lf = LineFormat()
                lf.opcode = XLSConstants.LINEFORMAT
                lf.data = lf.PROTOTYPE_BYTES
                lf.init()
                return lf
            }

        /**
         * get new Line Format in desired pattern and weight
         * <br></br>pattern= Pattern of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
         * <br></br>Weight of line	(-1= hairline, 0= narrow, 1= med (double), 2= wide (triple)
         * <br></br>note color is set to default
         *
         * @return new LineFormat record
         */
        fun getPrototype(style: Int, weight: Int): XLSRecord {
            val lf = LineFormat()
            lf.opcode = XLSConstants.LINEFORMAT
            lf.data = lf.PROTOTYPE_BYTES_1
            lf.init()
            lf.lineStyle = style
            lf.setLineWeight(weight)
            return lf
        }
    }
}
