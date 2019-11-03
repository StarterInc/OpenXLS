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
 * **Pos: Position Information(0x104f)**
 *
 *
 *
 *
 * for TextDisp, sets the label position as an offset from the default position
 * for PlotArea, used only for main axis + describes the plot-area bounding box; the tMainPlotArea in the SHTPROPS rec must be 1 or the POS rec is ignored
 * for Legend, describes legend pos + size
 *
 *
 * 4	rndTopLt		2		for PlotArea, TextDisp= 2; legend= 5 (what=3??? Data Table!)
 * 6	rndTopRt		2		for PlotArea, TextDisp= 2; legend: 1= use x2 and y2 for legend size; 2= autosize legend (ignore x2 + y2; if so, the fAutoSize bit of FRAME rec should be 1)
 * 8	x1				4		for PlotArea, x coord of bounding box; for TextDisp, horiz. offset from default pos; for legend, x coord in 1/4000
 * 12  y1				4		for PlotArea, y coord of bounding box; for TextDisp, vert. offset from default pos; for legend, y coord " "
 * 16	x2				4		for PlotArea, w of bounding box; for TextDisp, ignored; for legend= width
 * 20	y2				4		for PlotArea, h of bounding box; ""; for legend= height
 *
 *
 * Above is not correct;
 * Correct Information:
 * mdTopLt (2 bytes): A PositionMode structure that specifies the positioning mode for the upper-left corner of a legend,
 * an attached label, or the plot area. The valid combinations of mdTopLt and mdBotRt and the meaning of x1, y1, x2, y2
 * are specified in the Valid Combinations of mdTopLt and mdBotRt by Type table.
 * mdBotRt (2 bytes): A PositionMode structure that specifies the positioning mode for the lower-right corner of a legend,
 * an attached label, or the plot area. The valid combinations of mdTopLt and mdBotRt and the meaning of x1, y1, x2, y2
 * are specified in the following table.
 *
 *
 * x1 (2 bytes): A signed integer that specifies a position. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
 * unused1 (2 bytes): Undefined and MUST be ignored.
 * y1 (2 bytes): A signed integer that specifies a position. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
 * unused2 (2 bytes): Undefined and MUST be ignored.
 * x2 (2 bytes): A signed integer that specifies a width. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
 * unused3 (2 bytes): Undefined and MUST be ignored.
 * y2 (2 bytes): A signed integer that specifies a height. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
 * unused4 (2 bytes): Undefined and MUST be ignored.
 *
 *
 * Table:
 * Type			mdTopLtPosition Mode		mdBotRt Position Mode		Meaning
 * plot area (axis group)	MDPARENT			MDPARENT					The values of x1 and y1 specify the horizontal and vertical offsets of the primary axis group's
 * upper-left corner, relative to the upper-left corner of the chart area, in SPRC. The values of x2
 * and y2 specify the width and height of the primary axis group, in SPRC.
 * legend			MDCHART						MDABS						The values x1 and y1 specify the horizontal and vertical offsets of the legend's upper-left corner,
 * relative to the upper-left corner of the chart area, in SPRC. The values of x2 and y2 specify the
 * width and height of the legend, in points.
 * legend			MDCHART						MDPARENT					The values of x1 and y1 specify the horizontal and vertical offsets of the legend's upper-left corner,
 * relative to the upper-left corner of the chart area, in SPRC. The values of x2 and y2 MUST be ignored.
 * The size of the legend is determined by the application.
 * legend			MDKTH						MDPARENT					The values of x1, y1, x2 and y2 MUST be ignored. The legend is located inside a data table.
 *
 *
 * attached label	MDPARENT					MDPARENT					The meaning of x1 and y1 is specified in the Type of Attached Label table. x2 and y2 MUST be ignored.
 * The size of the attached label is determined by the application.
 *
 *
 *
 *
 * The PositionMode structure specifies positioning mode for position information saved in a Pos record.
 * Name			Value					   Meaning
 * MDFX			0x0000						Relative position to the chart, in points.
 * MDABS			0x0001						Absolute width and height in points. It can only be applied to the mdBotRt field of Pos.
 * MDPARENT		0x0002						Owner of Pos determines how to interpret the position data.
 * MDKTH			0x0003						Offset to default position, in 1/1000th of the plot area size.
 * MDCHART			0x0005						Relative position to the chart, in SPRC (A SPRC is a unit of measurement that is 1/4000th of the height or width of the chart).
 *
 *
 * Type of Attached Label	Meaning
 * Chart title				The value of x1 and y1 specify the horizontal and vertical offset of the title, relative to its default position, in SPRC.
 * Axis title				The value of x1 and y1 specify the offset of the title along the direction of a specific axis. The value of x1 specifies an offset along the category (3) axis, date axis, or horizontal value axis. The value of y1 specifies an offset along the value axis. Both offsets are relative to the title's default position, in 1/1000th of the axis length.
 * Data label				If the chart is not a pie chart group or a radar chart group, x1 and y1 specify the offset of the label along the direction of the specific axis.
 * The x1 value is an offset along the category (3) axis, date axis, or horizontal value axis.
 * The y1 value is an offset along the value axis, opposite to the direction of the value axis.
 * Both offsets are relative to the label's default position, in 1/1000th of the axis length.
 * For a pie chart group, the value of x1 specifies the clockwise angle, in degrees, and the value of y1 specifies the radius offset of the label relative to its default position, in 1/1000th of the pie radius length. A label moved toward the pie center has a negative radius offset.
 * For a radar chart group, the values of x1 and y1 specify the horizontal and vertical offset of the label relative to its default position, in 1/1000th of the axis length.
 */
class Pos : GenericChartObject(), ChartObject {
    internal var rndTopLt: Short = 0
    internal var rndTopRt: Short = 0
    internal var x1: Int = 0
    internal var y1: Int = 0
    internal var x2: Int = 0
    internal var y2: Int = 0

    private val PROTOTYPE_BYTES = byteArrayOf(2, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    /**
     * return the legend coordinates (x, y in 1/4000 of the chart height or width, w, h in points) or null, depending upon legend options
     * <br></br>NOTE: if the w or h are 0, use default.
     *
     * @return int[] x, y, w, h
     */
    /*The values x1 and y1 specify the horizontal and vertical offsets of the legend's upper-left corner,
    		relative to the upper-left corner of the chart area, in SPRC. The values of x2 and y2 specify the
		width and height of the legend, in points.*//*The values of x1 and y1 specify the horizontal and vertical offsets of the legend's upper-left corner,
		relative to the upper-left corner of the chart area, in SPRC. The values of x2 and y2 MUST be ignored.
		The size of the legend is determined by the application.*//*The values of x1, y1, x2 and y2 MUST be ignored. The legend is located inside a data table.*/ val legendCoords: IntArray?
        get() {
            if (rndTopLt.toInt() == 5 && rndTopRt.toInt() == 1) {
                return intArrayOf(x1, y1, x2, y2)
            } else if (rndTopLt.toInt() == 5 && rndTopRt.toInt() == 2) {
                return intArrayOf(x1, y1, 0, 0)
            } else if (rndTopLt.toInt() == 3 && rndTopRt.toInt() == 2) {
                return null
            }
            return null
        }

    /**
     * if plot area,
     * return the Plot Area (x, y, w, h in 1/4000 of the chart height or width) or null, depending upon options
     * <br></br>NOTES: Pos record of the Axis Parent:
     * <br></br>The Pos record specifies the position and size of the outer plot area. The outer plot area is the bounding rectangle that includes the axis labels,
     * <br></br>the axis titles, and data table of the chart.
     * if attached label,
     * Chart title			The value of x1 and y1 specify the horizontal and vertical offset of the title, relative to its default position, in SPRC.
     * Axis title			The value of x1 and y1 specify the offset of the title along the direction of a specific axis.
     * The value of x1 specifies an offset along the category axis, date axis, or horizontal value axis.
     * The value of y1 specifies an offset along the value axis. Both offsets are relative to the title's default position, in 1/1000th of the axis length.
     * Data label			If the chart is not a pie chart group or a radar chart group, x1 and y1 specify the offset of the label along the direction of the specific axis.
     * The x1 value is an offset along the category axis, date axis, or horizontal value axis. The y1 value is an offset along the value axis,
     * opposite to the direction of the value axis. Both offsets are relative to the label's default position, in 1/1000th of the axis length.
     * For a pie chart group, the value of x1 specifies the clockwise angle, in degrees, and the value of y1 specifies the radius offset of the label
     * relative to its default position, in 1/1000th of the pie radius length. A label moved toward the pie center has a negative radius offset.
     * For a radar chart group, the values of x1 and y1 specify the horizontal and vertical offset of the label relative to its default position,
     * in 1/1000th of the axis length.
     *
     * @return int[] x, y, w, h
     */
    /* 	The values of x1 and y1 specify the horizontal and vertical offsets of the primary axis group's
				upper-left corner, relative to the upper-left corner of the chart area, in SPRC. The values of x2
				and y2 specify the width and height of the primary axis group, in SPRC.*/ val coords: FloatArray?
        get() = if (rndTopLt.toInt() == 2 && rndTopRt.toInt() == 2) {
            floatArrayOf(x1.toFloat(), y1.toFloat(), x2.toFloat(), y2.toFloat())
        } else null

    override fun init() {
        super.init()
        // 0= in points, relative to the position of the chart
        // 1= absolute, in points
        // 2= parent of this rec determines
        // 3= offset to default
        // 5= relative, in 1/4000 of the w or h of the chart
        rndTopLt = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        rndTopRt = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        x1 = ByteTools.readInt(this.getBytesAt(4, 4)!!)
        y1 = ByteTools.readInt(this.getBytesAt(8, 4)!!)
        x2 = ByteTools.readInt(this.getBytesAt(12, 4)!!)
        y2 = ByteTools.readInt(this.getBytesAt(16, 4)!!)
    }

    /**
     * set the correct bytes for the desired type
     *
     * @param type
     */
    fun setType(type: Int) {
        when (type) {
            TYPE_PLOTAREA, TYPE_TEXTDISP -> {
                this.data[0] = 2
                this.data[1] = 0
            }
            TYPE_LEGEND -> {
                this.data[0] = 5
                this.data[1] = 2
            }
            TYPE_DATATABLE -> {
                this.data[0] = 3
                this.data[1] = 0
            }
        }
        this.data[2] = 2
        this.data[3] = 0
    }

    fun setX(x: Int) {
        x1 = x
        val b = ByteTools.cLongToLEBytes(x)
        System.arraycopy(b, 0, this.data!!, 4, 4)
    }

    fun setY(y: Int) {
        y1 = y
        val b = ByteTools.cLongToLEBytes(y)
        System.arraycopy(b, 0, this.data!!, 8, 4)
    }

    /**
     * for legends only, set width of bounds
     *
     * @param w
     */
    fun setLegendW(w: Int) {
        if (rndTopLt.toInt() == 5 && rndTopRt.toInt() == 1) {
            x2 = w
            val b = ByteTools.cLongToLEBytes(x2)
            System.arraycopy(b, 0, this.data!!, 12, 4)
        } // else throw exception?
    }

    /**
     * set autosize bit for legend Pos's
     * valid only for legend-type Pos's (rndTopLt==5)
     * also must set any associated Frame Autosize bits
     * [BugTracker 2844]
     */
    fun setAutosizeLegend() {
        if (this.rndTopLt.toInt() == 5) {// it's a legend Pos
            this.rndTopRt = 2// set to autoposition
            this.data[2] = 2
        }
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 7920967716354683818L
        val TYPE_TEXTDISP = 0
        val TYPE_LEGEND = 1
        val TYPE_PLOTAREA = 2
        val TYPE_DATATABLE = 3

        // TODO: Get def and parse options accordingly!
        fun getPrototype(type: Int): XLSRecord {
            val p = Pos()
            p.opcode = XLSConstants.POS
            p.data = p.PROTOTYPE_BYTES
            p.init()
            p.setType(type)
            return p
        }

        /**
         * convert a coordinate value in SPRC units to points
         *
         * @param val
         * @param w
         * @param h
         * @return
         */
        fun convertFromSPRC(`val`: Float, w: Float, h: Float): Float {
            // try this:
            return if (w != 0f) {
                (`val` / 4000.0).toFloat() * w
            } else {
                (`val` / 4000.0).toFloat() * h
            }
        }

        /**
         * convert a coordinate value in points to SPRC units
         * <br></br>Experimental at this point
         *
         * @param val
         * @param w
         * @param h
         * @return
         */
        fun convertToSPRC(`val`: Float, w: Float, h: Float): Float {
            // try this:
            return if (w != 0f) {
                (`val` * 4000.0).toFloat() / w
            } else {
                (`val` * 4000.0).toFloat() / h
            }
        }

        /**
         * convert a coordinate value in SPRC units to points
         *
         * @param val
         * @param w
         * @param h
         * @return
         */
        fun convertFromLabelUnits(`val`: Float, w: Float, h: Float): Float {
            // try this:
            return if (w != 0f) {
                (`val` / 1000.0).toFloat() * w
            } else {
                (`val` / 1000.0).toFloat() * h
            }
        }
    }
}
