/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or
 * modify
 * it under the terms of the GNU Lesser General Public
 * License as
 * published by the Free Software Foundation, either version
 * 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be
 * useful,
 * but WITHOUT ANY WARRANTY; without even the implied
 * warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
 * the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General
 * Public
 * License along with OpenXLS. If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */

package io.starter.formats.XLS.charts

import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools

/**
 * ** AttachedLabel: Series Data/Value Labels (0x100c) **
 *
 *
 *
 *
 * bit
 * 0	= show value label -- a bit that specifies whether the value, or the vertical value on bubble or scatter chart groups, is displayed in the data label.
 * 1   = show value as percentage -- (must be 0 for non-pie)
 * 2	= show cat/label as percentage -- A bit that specifies whether the category (3) name and value,
 * represented as a percentage of the sum of the values of the series with
 * which the data label is associated, are displayed in the data label.	(pie charts only)
 * 3	= unused
 * 4	= show cat or series label -- A bit that specifies whether the category, or the horizontal value on bubble or scatter chart groups,
 * is displayed in the data label on a non-area chart group, or the series name is displayed in the data label on an area chart group.
 * 5	= show bubble label	-- A bit that specifies whether the bubble size is displayed in the data label.
 * 6   = show series name -- A bit that specifies whether the data label contains the name of the series.
 *
 *
 * ???? Series DataLabels do not appear to be done via AttachedLabel
 */
class AttachedLabel : GenericChartObject(), ChartObject {

    private var grbit: Short = 0

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0)

    /**
     * return a string of all the label options chosen
     *
     * @return
     */
    // bit 0
    // bit 1
    // bit 2
    // Pie only
    // bit 4
    // bit 5
    // bit 6
    val type: String
        get() {
            var ret = ""
            if (grbit and VALUE == VALUE)
                ret = "Value "
            if (grbit and VALUEPERCENT == VALUEPERCENT)
                ret = ret + "ValuePerecentage "
            if (grbit and CATEGORYPERCENT == CATEGORYPERCENT)
                ret = ret + "CategoryPercentage "
            if (grbit and CATEGORYLABEL == CATEGORYLABEL)
                ret = ret + "CategoryLabel "
            if (grbit and BUBBLELABEL == BUBBLELABEL)
                ret = ret + "BubbleLabel "
            if (grbit and VALUELABEL == VALUELABEL)
                ret = ret + "SeriesLabel "
            return ret.trim { it <= ' ' }
        }

    /**
     * return the data label options as an int
     *
     * @return a combination of data label options above or 0 if none
     * @see AttachedLabel constants
     * <br></br>SHOWVALUE= 0x1;
     * <br></br>SHOWVALUEPERCENT= 0x2;
     * <br></br>SHOWCATEGORYPERCENT= 0x4;
     * <br></br>SHOWCATEGORYLABEL= 0x10;
     * <br></br>SHOWBUBBLELABEL= 0x20;
     * <br></br>SHOWSERIESLABEL= 0x40;
     */
    val typeInt: Int
        get() = grbit.toInt()

    override fun init() {
        super.init()
        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
    }

    fun setType(type: Short) {
        grbit = type
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    /**
     * Show or Hide AttachedLabel Option
     * <br></br>ShowValueLabel,
     * <br></br>ShowValueAsPercent,
     * <br></br>ShowLabelAsPercent,
     * <br></br>ShowLabel,
     * <br></br>ShowBubbleLabel
     * <br></br>ShowSeriesName
     *
     * @param val true or 1 to set
     */
    fun setType(type: String, `val`: String) {
        val bSet = `val` == "true" || `val` == "1"
        if (type == "ShowValueLabel")
            grbit = ByteTools.updateGrBit(grbit, bSet, 0)
        if (type == "ShowValueAsPercent")
            grbit = ByteTools.updateGrBit(grbit, bSet, 1)
        if (type == "ShowLabelAsPercent")
            grbit = ByteTools.updateGrBit(grbit, bSet, 2)
        if (type == "ShowLabel")
            grbit = ByteTools.updateGrBit(grbit, bSet, 4)
        if (type == "ShowBubbleLabel")
            grbit = ByteTools.updateGrBit(grbit, bSet, 5)
        if (type == "ShowSeriesName")
            grbit = ByteTools.updateGrBit(grbit, bSet, 6)
        val bb = ByteTools.shortToLEBytes(grbit)
        this.data[0] = bb[0]
        this.data[1] = bb[1]
    }

    /**
     * return the value of the specified option
     *
     * @param type Sting option
     * @return string true or false
     */
    fun getType(type: String): String {
        var b = false

        if (type == "ShowValueLabel")
            b = grbit and VALUE == VALUE // bit 0
        if (type == "ShowValueAsPercent")
            b = grbit and VALUEPERCENT == VALUEPERCENT
        if (type == "ShowLabelAsPercent")
            b = grbit and CATEGORYPERCENT == CATEGORYPERCENT
        if (type == "ShowLabel")
            b = grbit and CATEGORYLABEL == CATEGORYLABEL
        if (type == "ShowBubbleLabel")
            b = grbit and BUBBLELABEL == BUBBLELABEL
        if (type == "ShowSeriesName")
            b = grbit and VALUELABEL == VALUELABEL
        return if (b) "1" else "0"
    }

    /**
     * @param type
     */
    @Deprecated("")
    fun setType(type: String) {
        var t: Short = 0
        if (type.equals("Value", ignoreCase = true) || type.equals("Y Value", ignoreCase = true))
            t = 1
        else if (type.equals("ValuePercentage", ignoreCase = true))
            t = 2
        else if (type.equals("CategoryPercentage", ignoreCase = true))
            t = 3
        else if (type.equals("Category", ignoreCase = true) || type.equals("X Value", ignoreCase = true))
            t = 16
        else if (type.equals("CandP", ignoreCase = true))
            t = 22
        else if (type.equals("Bubble", ignoreCase = true))
            t = 32
        grbit = t
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    override fun toString(): String {
        return this.type
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 2532517522176536995L
        val VALUE = 0x1
        val VALUEPERCENT = 0x2
        val CATEGORYPERCENT = 0x4
        val CATEGORYLABEL = 0x10
        val BUBBLELABEL = 0x20
        val VALUELABEL = 0x40

        val prototype: XLSRecord?
            get() {
                val al = AttachedLabel()
                al.opcode = XLSConstants.ATTACHEDLABEL
                al.data = al.PROTOTYPE_BYTES
                al.init()
                return al
            }
    }

}
