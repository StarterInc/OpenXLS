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
 * **CatserRange: Defines a Category or Series Axis (0x1020)**
 *
 *
 * 4		catCross	2	Value axis/category crossing point (2-D charts only)
 * If fMaxCross is set to 1, the value this field MUST be ignored.
 * for series axes, must be 0
 * for cat: This field specifies the category at which the value axis crosses.
 * For example, if this field is 2, the value axis crosses this axis at the second category on this axis.
 * MUST be greater than or equal to 1 and less than or equal to 31999.
 * 6		catLabel	2	Frequency of labels. A signed integer that specifies the interval between axis labels on this axis.
 * MUST be greater than or equal to 1 and less than or equal to 31999.
 * MUST be ignored for a date axis.
 * 8		catMark		2	Frequency of tick marks.  A signed integer that specifies the interval at which major tick marks
 * and minor tick marks are displayed on the axis. Major tick marks and minor tick marks that would
 * have been visible are hidden unless they are located at a multiple of this field.
 * MUST be greater than or equal to 1, and less than or equal to 31999.
 * MUST be ignored for a date axis.
 * 10		grbit		2	Format flags
 *
 *
 * The catCross field defines the point on the category axis where the value axis crosses.
 * A value of 01 indicates that the value axis crosses to the left, or in the center, of the first category
 * (depending on the value of bit 0 of the grbit field); a value of 02 indicates that the value axis crosses
 * to the left or center of the second category, and so on. Bit 2 of the grbit field overrides the value of catCross when set to 1.
 *
 *
 * The catLabel field defines how often labels appear along the category or series axis. A value of 01 indicates
 * that a category label will appear with each category, a value of 02 means a label appears every other category, and so on.
 *
 *
 * The catMark field defines how often tick marks appear along the category or series axis. A value of 01 indicates
 * that a tick mark will appear between each category or series; a value of 02 means a label appears between every
 * other category or series, etc.
 *
 *
 * format flags:
 * 0	0xOl		fBetween		Value axis crossing a = axis crosses midcategory I = axis crosses between categories
 * 1	0x02		fMaxCross		Value axis crosses at the far right category (in a line, bar, column, scatter, or area chart; 2-D charts only)
 * 0 The value axis crosses this axis at the value specified by catCross.
 * 1 The value axis crosses this axis at the last category, the last series, or the maximum date.
 * 2	0x04		fReverse		Display categories in reverse order
 * 7-3	0xF8		(reserved)		Reserved; must be zero
 * 7-0	0xFF		(reserved)		Reserved; must be zero
 */
class CatserRange : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    private var catCross: Short = 0
    private var catLabel: Short = 0
    private var catMark: Short = 0
    private var fBetween: Boolean = false
    private var fMaxCross: Boolean = false
    /**
     * returns true if the axis should be displayed at top of chart area
     * false if axis is displayed in the default bottom location
     *
     * @return
     */
    var isReversed: Boolean = false
        private set

    private val PROTOTYPE_BYTES = byteArrayOf(1, 0, 1, 0, 1, 0, 1, 0)

    var crossBetween: Boolean
        get() = fBetween
        set(b) {
            fBetween = b
            grbit = ByteTools.updateGrBit(grbit, fBetween, 0)
            updateRecord()
        }

    var crossMax: Boolean
        get() = fMaxCross
        set(b) {
            fMaxCross = b
            grbit = ByteTools.updateGrBit(grbit, fMaxCross, 1)
            updateRecord()
        }

    // 20070727 KSC: TODO: Get data def and parse correctly!
    fun setOpt(op: Int) {
        val b = ByteTools.shortToLEBytes(op.toShort())
        this.data[0] = b[0]
        this.data[1] = b[1]
        // 20070802 KSC: don't know what this means
        this.data[6] = 0
    }

    // 20070802 KSC: parse data
    override fun init() {
        super.init()
        catCross = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        catLabel = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        catMark = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        grbit = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        fBetween = grbit and 0x1 == 0x1
        fMaxCross = grbit and 0x2 == 0x2
        isReversed = grbit and 0x4 == 0x4
    }

    private fun updateRecord() {
        var b = ByteTools.shortToLEBytes(catCross)
        this.data[0] = b[0]
        this.data[1] = b[1]
        b = ByteTools.shortToLEBytes(catLabel)
        this.data[2] = b[0]
        this.data[3] = b[1]
        b = ByteTools.shortToLEBytes(catMark)
        this.data[4] = b[0]
        this.data[5] = b[1]
        b = ByteTools.shortToLEBytes(grbit)
        this.data[6] = b[0]
        this.data[7] = b[1]
    }

    // get/set methods
    fun getCatCross(): Int {
        return catCross.toInt()
    }

    fun getCatLabel(): Int {
        return catLabel.toInt()
    }

    fun getCatMark(): Int {
        return catMark.toInt()
    }

    fun setCatCross(c: Int) {
        catCross = c.toShort()
        updateRecord()
    }

    fun setCatLabel(c: Int) {
        catLabel = c.toShort()
        updateRecord()
    }

    fun setCatMark(c: Int) {
        catMark = c.toShort()
        updateRecord()
    }

    /**
     * sets a specific OOXML axis option
     * <br></br>can be one of:
     * <br></br>auto
     * <br></br>crosses			possible crossing points (autoZero, max, min)
     * <br></br>crossesAt		where on axis the perpendicular axis crosses (double val)
     * <br></br>lblAlign			text alignment for tick labels (ctr, l, r) (cat only)
     * <br></br>lblOffset		distance of labels from the axis (0-1000)  (cat only)
     * <br></br>tickLblSkip		how many tick labels to skip between label (int >= 1)
     * <br></br>tickMarkSkip		how many tick marks to skip betwen ticks (int >= 1)
     *
     * @param op
     * @param val
     */
    // TODO: auto
    fun setOption(op: String, `val`: String): Boolean {
        if (op == "crossesAt")
        // specifies where axis crosses (double value
            catCross = java.lang.Short.valueOf(`val`).toShort()
        else if (op == "orientation") {    // axis orientation minMax or maxMin  -- fReverse
            isReversed = `val` == "maxMin"    // means in reverse order
            ByteTools.updateGrBit(grbit, isReversed, 2)
        } else if (op == "crosses") {            // specifies how axis crosses it's perpendicular axis (val= max, min, autoZero)  -- fbetween + fMaxCross?/fAutoCross + fMaxCross
            if (`val` == "max") {        // TODO: this is probly wrong
                fMaxCross = true
                ByteTools.updateGrBit(grbit, fMaxCross, 7)
            } else if (`val` == "mid") {
                fBetween = false
                ByteTools.updateGrBit(grbit, fBetween, 0)
            } else if (`val` == "autoZero") {
                fBetween = true    // is this correct??
                ByteTools.updateGrBit(grbit, fBetween, 0)
            } else if (`val` == "min");
            // TODO:  ???
        } else if (op == "tickMarkSkip")
        //val= how many tick marks to skip before next one is drawn -- catMark -- catsterrange only?
            catMark = Integer.valueOf(`val`).toShort()
        else if (op == "tickLblSkip")
            catLabel = Integer.valueOf(`val`).toShort()
        else
            return false    // not handled
        this.updateRecord()
        return true
    }


    /**
     * retrieve generic Category axis option
     *
     * @param op
     * @return
     */
    fun getOption(op: String): String? {
        // TODO: auto, lblAlign, lblOffset
        if (op == "crossesAt")
            return catCross.toString()
        if (op == "orientation")
            return if (isReversed) "maxMin" else "minMax"
        if (op == "crosses") {
            if (fMaxCross) return "max"
            return if (fBetween) "autoZero" else "min"    // correct??
// correct??
        }
        if (op == "tickMarkSkip")
        //val= how many tick marks to skip before next one is drawn -- catMark -- catsterrange only?
            return catMark.toString()
        return if (op == "tickLblSkip") catLabel.toString() else null
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 905038625844435651L

        // 20070723 KSC: Need to create new records
        val prototype: XLSRecord?
            get() {
                val c = CatserRange()
                c.opcode = XLSConstants.CATSERRANGE
                c.data = c.PROTOTYPE_BYTES
                c.init()
                return c
            }    // 0, 0, 1, 0, 1, 0, 0, 0	- for axis 2 of surface chart
    }
}
