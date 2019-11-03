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

import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.OOXML.SpPr
import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.Stack

/**
 * **ChartLine: Drop/Hi-Lo/Series Lines on a Line Chart (0x101c)**
 *
 *
 * The CrtLine record specifies the presence of drop lines, high-low lines, series lines or leader lines on the chart group.
 * This record is followed by a LineFormat record which specifies the format of the lines.
 *
 * <br></br> id (2 bytes): An unsigned integer that specifies the type of line that is present on the chart group.
 * This field value MUST be unique among the other id field values in CrtLine records in the current chart group.
 * This field MUST be greater than the id field values in preceding CrtLine records in the current chart group. MUST be a value from the following table:
 *
 *
 * Value		Type of Line
 * 0x0000		Drop lines below the data points of line, area, and stock chart groups.
 * 0x0001		High-low lines around the data points of line and stock chart groups.
 * 0x0002		Series lines connecting data points of stacked column and bar chart groups, and the primary pie to the secondary bar/pie of bar of pie and pie of pie chart groups.
 * 0x0003		Leader lines with non-default formatting connecting data labels to the data point of pie and pie of pie chart groups.
 *
 *
 * But also there is:
 * DROPBAR = DropBar Begin LineFormat AreaFormat [GELFRAME] [SHAPEPROPS] End
 */
class ChartLine : GenericChartObject(), ChartObject {
    private var id: Short = 0
    private var cf: ChartFormat? = null    // necessary to link to corresponding ChartFormat

    /**
     * <br></br>0= drop lines below the data points of line, area and stock charts
     * <br></br>1= High-low lines around the data points of line and stock charts
     * <br></br>2- Series Line connecting data points of stacked column and bar charts + some pie chart configurations
     * <br></br>3= Leader lines with non-default formatting for pie and pie of pie
     *
     * @return
     */
    /**
     * sets the chart line type:
     * <br></br>0= drop lines below the data points of line, area and stock charts
     * <br></br>1= High-low lines around the data points of line and stock charts
     * <br></br>2- Series Line connecting data points of stacked column and bar charts + some pie chart configurations
     * <br></br>3= Leader lines with non-default formatting for pie and pie of pie
     */
    var lineType: Int
        get() = id.toInt()
        set(id) {
            this.id = id.toShort()
            val b = ByteTools.shortToLEBytes(this.id)
            this.data[0] = b[0]
            this.data[1] = b[1]
        }

    /**
     * return the OOXML to define this ChartLine
     *
     * @return
     */
    val ooxml: StringBuffer
        get() {
            val cooxml = StringBuffer()
            var tag: String? = null
            if (id == TYPE_DROPLINE.toShort())
                tag = "c:dropLines>"
            else if (id == TYPE_HILOWLINE.toShort())
                tag = "c:hiLowLines>"
            else if (id == TYPE_SERIESLINE.toShort())
                tag = "c:serLines>"
            cooxml.append("<" + tag!!)
            val lf = findLineFormatRec()
            if (lf != null)
                cooxml.append(lf.ooxml)
            cooxml.append("</$tag")
            return cooxml
        }

    override fun init() {
        super.init()
        id = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
    }

    /**
     * return the LineFormat rec associated with this ChartLine
     * (== the next record in the cf chart array)
     *
     * @return
     */
    private fun findLineFormatRec(): LineFormat? {
        var lf: LineFormat? = null
        for (i in cf!!.chartArr.indices) {
            if (cf!!.chartArr[i] == this) {
                lf = cf!!.chartArr[i + 1] as LineFormat
                break
            }
        }
        return lf
    }

    /**
     * parse a Chart Line OOXML element: either
     *  * dropLines
     *  * hiLowLines
     *  * serLines
     *
     * @param xpp
     * @param lastTag
     * @param cf
     */
    fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, cf: ChartFormat, bk: WorkBookHandle) {
        this.cf = cf
        val endTag = lastTag.peek()
        try {
            var eventType = xpp.eventType
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    val tnm = xpp.name
                    if (tnm == "spPr") {
                        lastTag.push(tnm)
                        val sppr = SpPr.parseOOXML(xpp, lastTag, bk).cloneElement() as SpPr
                        val lf = findLineFormatRec()
                        lf?.setFromOOXML(sppr)
                    }
                } else if (eventType == XmlPullParser.END_TAG) {
                    if (xpp.name == endTag)
                        lastTag.pop()
                    break
                }
                eventType = xpp.next()
            }
        } catch (e: Exception) {
            Logger.logErr("ChartLine.parseOOXML: $e")
        }

    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -8311605814020380069L
        var TYPE_DROPLINE: Byte = 0
        var TYPE_HILOWLINE: Byte = 1
        var TYPE_SERIESLINE: Byte = 2
        var TYPE_LEADERLINE: Byte = 3

        val prototype: XLSRecord?
            get() {
                val cl = ChartLine()
                cl.opcode = XLSConstants.CHARTLINE
                cl.data = byteArrayOf(0, 0)
                cl.init()
                return cl
            }
    }
}
