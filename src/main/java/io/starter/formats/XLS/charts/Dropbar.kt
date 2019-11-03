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

import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools

/**
 * **Dropbar: Defines Drop Bars (0x103d)**
 * Controls up or down bars on a line (or stock, for 2007 v) chart with multiple series
 * the first dropBar record controls upBars
 * the second record controls downBars
 * Also, if these records exist, SeriesList cSer > 1
 *
 *
 *
 *
 *
 *
 * pcGap (2 bytes): A signed integer that specifies the width of the gap between the up bars or the down bars. MUST be a value between 0 and 500.
 * The width of the gap in SPRCs can be calculated by the following formula:
 * Width of the gap in SPRCs = 1 + pcGap
 *
 * <br></br>
 * The DropBar record occurs in the ChartFormat subrecord stream after the Legend record,
 * and contains subrecords LineFormat, AreaFormat, [GelFrame], [ShapeProps]
 */
class Dropbar : GenericChartObject(), ChartObject {
    /**
     * returns the width of the gap between the up bars or the down bars.
     * <br></br>The gap width is a value between 0 and 500.
     * <br></br>The width of the gap in SPRCs can be calculated by the following formula:
     * Width of the gap in SPRCs = 1 + pcGap
     *
     * @return
     */
    var gapWidth: Short = 0
        internal set

    override fun init() {
        super.init()
        gapWidth = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
    }

    /**
     * sets the width of the gap between the up bars or the down bars. MUST be a value between 0 and 500.
     * The width of the gap in SPRCs can be calculated by the following formula:
     * Width of the gap in SPRCs = 1 + pcGap
     *
     * @param gap
     */
    fun setGapWidth(gap: Int) {
        gapWidth = gap.toShort()
        val b = ByteTools.shortToLEBytes(gapWidth)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    /**
     * return the OOXML to define this ChartLine
     *
     * @return
     */
    fun getOOXML(upBars: Boolean): StringBuffer {
        val cooxml = StringBuffer()
        val tag: String
        if (upBars)
            tag = "c:upBars>"
        else
            tag = "c:downBars>"
        cooxml.append("<$tag")
        val lf = chartArr[0] as LineFormat
        cooxml.append(lf.ooxml)
        // TODO: finish this logic
        if (!parentChart!!.workBook!!.isExcel2007) {
            val af = chartArr[1] as AreaFormat
            cooxml.append(af.ooxml)
        }
        cooxml.append("</$tag")
        return cooxml
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 6826327230442065566L

        // 150 is default gap width
        val prototype: XLSRecord?
            get() {
                val db = Dropbar()
                db.opcode = XLSConstants.DROPBAR
                db.data = byteArrayOf(-106, 0)
                db.init()
                return db
            }
    }
}
