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

/**
 * CrtLayout12A		0x08A7
 * The CrtLayout12A record specifies layout information for a plot area.
 *
 *
 * frtHeader (12 bytes): An FrtHeader structure. The frtheader.rt field MUST be 0x08A7.
 *
 *
 * dwCheckSum (4 bytes): An unsigned integer that specifies the checksum. MUST be a value from the following table:
 * fManPlotArea field of ShtProps		fAlwaysAutoPlotArea field of ShtProps		dwCheckSum
 * 0x0									0x0											0x00000001
 * 0x0									0x1											0x00000000
 * 0x1									0x0											0x00000000
 * 0x1									0x1											0x00000001
 *
 *
 * A - fLayoutTargetInner (1 bit): A bit that specifies the type of plot area for the layout target.
 * Value		Meaning
 * 0x0			Outer plot area - The bounding rectangle that includes the axis labels, axis titles, data table (2) and plot area of the chart.
 * 0x1			Inner plot area – The rectangle bounded by the chart axes.
 *
 *
 * reserved1 (15 bits): MUST be zero, and MUST be ignored.
 *
 *
 * xTL (2 bytes): A signed integer that specifies the horizontal offset of the plot area’s upper-left corner, relative to the upper-left corner of the chart area, in SPRC.
 * yTL (2 bytes): A signed integer that specifies the vertical offset of the plot area’s upper-left corner, relative to the upper-left corner of the chart area, in SPRC.
 * xBR (2 bytes): A signed integer that specifies the width of the plot area, in SPRC.
 * yBR (2 bytes): A signed integer that specifies the height of the plot area, in SPRC.
 * wXMode (2 bytes): A CrtLayout12Mode structure that specifies the meaning of x.
 * wYMode (2 bytes): A CrtLayout12Mode structure that specifies the meaning of y.
 * wWidthMode (2 bytes): A CrtLayout12Mode structure that specifies the meaning of dx.
 * wHeightMode (2 bytes): A CrtLayout12Mode structure that specifies the meaning of dy.
 * x (8 bytes): An Xnum value that specifies a horizontal offset. The meaning is determined by wXMode.
 * y (8 bytes): An Xnum value that specifies a vertical offset. The meaning is determined by wYMode.
 * dx (8 bytes): An Xnum value that specifies a width or a horizontal offset. The meaning is determined by wWidthMode.
 * dy (8 bytes): An Xnum value that specifies a height or a vertical offset. The meaning is determined by wHeightMode.
 * reserved2 (2 bytes): MUST be zero, and MUST be ignored.
 *
 *
 * The CrtLayout12Mode record specifies a layout mode. Each layout mode specifies a different meaning of the x, y, dx, and dy fields of CrtLayout12 and CrtLayout12A.
 * Name		Value		Meaning
 * L12MAUTO	0x0000		Position and dimension (2) are determined by the application. x, y, dx and dy MUST be ignored.
 * L12MFACTOR	0x0001		x and y specify the offset of the top left corner, relative to its default position, as a fraction of the chart area. MUST be greater than or equal to -1.0 and MUST be less than or equal to 1.0. dx and dy specify the width and height, as a fraction of the chart area, MUST be greater than or equal to 0.0, and MUST be less than or equal to 1.0.
 * L12MEDGE	0x0002		x and y specify the offset of the upper-left corner; dx and dy specify the offset of the bottom-right corner. x, y, dx and dy are specified relative to the upper-left corner of the chart area as a fraction of the chart area. x, y, dx and dy MUST be greater than or equal to 0.0, and MUST be less than or equal to 1.0
 *
 *
 * Xnum is a 64-bit binary floating-point number as specified in [IEEE754]. This value MUST NOT be infinity, denormalized, not-a-number (NaN), nor negative zero.
 *
 *
 * A SPRC is a unit of measurement that is 1/4000th of the height or width of the chart. If the field is being used to specify a width or horizontal distance, the SPRC is 1/4000th of the width of the chart. If the field is being used to specify a height or vertical distance, the SPRC is 1/4000th of the height of the chart.
 */

import io.starter.toolkit.ByteTools

class CrtLayout12A : GenericChartObject(), ChartObject {
    internal var fLayoutTargetInner: Byte = 0
    internal var xTL: Short = 0
    internal var yTL: Short = 0
    internal var xBR: Short = 0
    internal var yBR: Short = 0
    internal var wXMode: Short = 0
    internal var wYMode: Short = 0
    internal var wWidthMode: Short = 0
    internal var wHeightMode: Short = 0
    internal var x: Float = 0.toFloat()
    internal var y: Float = 0.toFloat()
    internal var dx: Float = 0.toFloat()
    internal var dy: Float = 0.toFloat()

    override fun init() {
        super.init()
        val data = this.data
        fLayoutTargetInner = data!![16]
        xTL = ByteTools.readShort(data[18].toInt(), data[19].toInt())
        yTL = ByteTools.readShort(data[20].toInt(), data[21].toInt())
        xBR = ByteTools.readShort(data[22].toInt(), data[23].toInt())
        yBR = ByteTools.readShort(data[24].toInt(), data[25].toInt())
        wXMode = ByteTools.readShort(data[26].toInt(), data[27].toInt())
        wYMode = ByteTools.readShort(data[28].toInt(), data[29].toInt())
        wWidthMode = ByteTools.readShort(data[30].toInt(), data[31].toInt())
        wHeightMode = ByteTools.readShort(data[32].toInt(), data[33].toInt())
        x = ByteTools.eightBytetoLEDouble(this.getBytesAt(34, 8)!!).toFloat()
        y = ByteTools.eightBytetoLEDouble(this.getBytesAt(42, 8)!!).toFloat()
        dx = ByteTools.eightBytetoLEDouble(this.getBytesAt(50, 8)!!).toFloat()
        dy = ByteTools.eightBytetoLEDouble(this.getBytesAt(58, 8)!!).toFloat()
    }

    /**
     * If this CrtLayout contains information for the inner plot area coords
     * and the coordinates are NOT determined by the application,
     * calculate the plot area coordinates.
     * <br></br>If not possible, return null;
     * <br></br>[Inner plot area – The rectangle bounded by the chart axes]
     * @param w            chart width for SPRC calc
     * @param h            chart height for SPRC calc
     * @return
     */
    internal fun getInnerPlotCoords(w: Float, h: Float): FloatArray? {
        /*L12MAUTO	0x0000		Position and dimension (2) are determined by the application. x, y, dx and dy MUST be ignored.
		L12MFACTOR	0x0001		x and y specify the offset of the top left corner, relative to its default position, as a fraction of the chart area. MUST be greater than or equal to -1.0 and MUST be less than or equal to 1.0. dx and dy specify the width and height, as a fraction of the chart area, MUST be greater than or equal to 0.0, and MUST be less than or equal to 1.0.
		L12MEDGE	0x0002		x and y specify the offset of the upper-left corner; dx and dy specify the offset of the bottom-right corner. x, y, dx and dy are specified relative to the upper-left corner of the chart area as a fraction of the chart area. x, y, dx and dy MUST be greater than or equal to 0.0, and MUST be less than or equal to 1.0
		*/
        return if (fLayoutTargetInner.toInt() == 0) null else floatArrayOf(Pos.convertFromSPRC(xTL.toFloat(), w, 0f), // x offset
                Pos.convertFromSPRC(yTL.toFloat(), 0f, h), // y offset
                Pos.convertFromSPRC(xBR.toFloat(), w, 0f), // w
                Pos.convertFromSPRC(yBR.toFloat(), 0f, h))
// h
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -1868700214505277636L
    }

}
