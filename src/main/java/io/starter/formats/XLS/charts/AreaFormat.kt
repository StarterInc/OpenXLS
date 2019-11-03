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
import io.starter.formats.OOXML.FillGroup
import io.starter.formats.OOXML.SpPr
import io.starter.formats.XLS.FormatConstants
import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools

/**
 * **AreaFormat: Colors and Patterns for an area(0x100a)**
 *
 *
 * 4	rgbFore		4		FgColor RGB Value (High byte is 0)
 * 8	rgbBack		4		Bg Color ""
 * 12	fls			2		Pattern
 * 14	grbit		2		flags
 * 16	icvFore		2 		index to fg color
 * 18	icvBack		2		index to bg color
 *
 *
 * grbit:
 * 0		0x1		fAuto		Automatic Format
 * 1		0x2		fInvertNeg	Fg + Bg is swapped when data is neg.
 */

/**
 * more info on colors:
 *
 * The chart color table is a subset of the full color table.
 * icv (2 bytes): An Icv that specifies a color from the chart color table.
 * MUST be greater than or equal to 0x0008 and less than or equal to 0x003F, or greater than or equal to 0x004D and less than or equal to 0x004F.
 *
 * This info is not yet verified:
 *
 * For icvFore, icvBack, must be either 0-7 or 0x40 or 0x41 or icv????
 * "The default value of this field is selected automatically from the next available color in the Chart color table."
 *
 * For icvBack, must be either 0-7 or 0x40 or 0x41
 * The default value of this field is 0x0009.
 *
 * icv (chart color table index):  > 0x0008 < 003F OR 0x4D, 0x4E, 0x4F OR 0x7FFF
 *
 * 0-7 are bascially the normal COLORTABLE entries
 * The icv values greater than or equal to 0x0008 and less than or equal to 0x003F, specify the palette colors in the table.
 * If a Palette record exists in this file, these icv values specify colors from the rgColor array in the Palette record.
 * If no Palette record exists, these values specify colors in the default palette.
 *
 * The next 56 values in this part of the color table are specified as follows:
 * 0x0040
 * Default foreground color. This is the window text color in the data sheet display.
 * 0x0041
 * Default background color. This is the window background color in the data sheet display and is the default background color for a cell.
 * 0x004D
 * Default chart foreground color. This is the window text color in the chart display.
 * 0x004E
 * Default chart background color. This is the window background color in the chart display.
 * 0x004F
 * Chart neutral color which is black, an RGB value of (0,0,0).
 * 0x7FFF
 * Font automatic color. This is the window text color.
 */
class AreaFormat : GenericChartObject(), ChartObject {
    private var rgbFore: java.awt.Color? = null
    private var rgbBack: java.awt.Color? = null
    private var fls: Short = 0
    private var grbit: Short = 0
    private var icvFore: Short = 0
    private var icvBack: Short = 0
    internal var fAuto: Boolean = false
    internal var fInvertNeg: Boolean = false

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 77, 0, 77, 0)
    private val PROTOTYPE_BYTES_1 = byteArrayOf(-64, -64, -64, 0, 0, 0, 0, 0, 1, 0, 0, 0, 22, 0, 79, 0)
    private val PROTOTYPE_BYTES_2 = byteArrayOf(0, 0, 0, 0, -1, -1, -1, 0, 1, 0, 1, 0, 77, 0, 78, 0)

    /**
     * return the area fill color
     * @return int index into color table
     */
    // no fill
    // solid; forecolor is bg
    // medium grey
    // dark grey (possible for more options not yet handled
    // light grey
    // rest are actual fill patterns TODO handle
    val fillColor: Int
        get() {
            if (fls.toInt() == 0)
                return FormatConstants.COLOR_WHITE
            if (fls.toInt() == 1)
                return geticvFore()
            if (fls.toInt() == 2)
                return FormatConstants.COLOR_GRAY50
            if (fls.toInt() == 3)
                return FormatConstants.COLOR_GRAY80
            return if (fls.toInt() == 4) FormatConstants.COLOR_GRAY25 else geticvFore()
        }

    /**
     * return the area fill color
     * @return String color hex string
     */
    // no fill
    // solid; forecolor is bg
    // medium grey
    // dark grey (possible for more options not yet handled
    // light grey
    // rest are actual fill patterns TODO handle
    val fillColorStr: String?
        get() {
            if (fAuto)
                return null
            if (fls.toInt() == 0)
                return "#FFFFFF"
            if (fls.toInt() == 1)
                return FormatHandle.colorToHexString(rgbFore!!)
            if (fls.toInt() == 2)
                return FormatHandle.colorToHexString(FormatConstants.Gray50)
            if (fls.toInt() == 3)
                return FormatHandle.colorToHexString(FormatConstants.Gray80)
            return if (fls.toInt() == 4) FormatHandle.colorToHexString(FormatConstants.Gray25) else FormatHandle.colorToHexString(rgbFore!!)
        }

    //    	ooxml.append("<a:srgbClr val=\"" + FormatHandle.colorToHexString(rgb) + "\"/>");
    val ooxml: StringBuffer
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:solidFill>")
            ooxml.append("</a:solidFill>")
            return ooxml
        }

    override fun init() {
        super.init()
        val data = this.data
        grbit = ByteTools.readShort(data!![10].toInt(), data[11].toInt())
        fAuto = grbit and 0x1 == 0x1
        fInvertNeg = grbit and 0x2 == 0x2
        rgbFore = java.awt.Color(if (data[0] < 0) 255 + data[0] else data[0], if (data[1] < 0) 255 + data[1] else data[1], if (data[2] < 0) 255 + data[2] else data[2])
        rgbBack = java.awt.Color(if (data[4] < 0) 255 + data[4] else data[4], if (data[5] < 0) 255 + data[5] else data[5], if (data[6] < 0) 255 + data[6] else data[6])
        fls = ByteTools.readShort(data[8].toInt(), data[9].toInt())
        icvFore = ByteTools.readShort(data[12].toInt(), data[13].toInt())
        icvBack = ByteTools.readShort(data[14].toInt(), data[15].toInt())
    }


    private fun updateRecord() {
        var b = ByteArray(4)
        b[0] = rgbFore!!.red.toByte()
        b[1] = rgbFore!!.green.toByte()
        b[2] = rgbFore!!.blue.toByte()
        b[3] = 0    // reserved/0
        System.arraycopy(b, 0, this.data!!, 0, 4)
        b[0] = rgbBack!!.red.toByte()
        b[1] = rgbBack!!.green.toByte()
        b[2] = rgbBack!!.blue.toByte()
        b[3] = 0    // reserved/0
        System.arraycopy(b, 0, this.data!!, 4, 4)
        b = ByteTools.shortToLEBytes(fls)
        this.data[8] = b[0]
        this.data[9] = b[1]
        grbit = 0
        if (fAuto) grbit = 0x1
        if (fInvertNeg) grbit = grbit or 0x2
        b = ByteTools.shortToLEBytes(grbit)
        this.data[10] = b[0]
        this.data[11] = b[1]
        b = ByteTools.shortToLEBytes(icvFore)
        this.data[12] = b[0]
        this.data[13] = b[1]
        b = ByteTools.shortToLEBytes(icvBack)
        this.data[14] = b[0]
        this.data[15] = b[1]
    }


    override fun toString(): String {
        return "AreaFormat: Pattern=" + fls + " ForeColor=" + icvFore + " BackColor=" + icvBack + " Automatic Format=" + (grbit and 0x1 == 0x1)
    }

    /**
     * return the bg color index for this Area Format
     * @return
     */
    fun geticvBack(): Int {
        return if (icvBack > this.colorTable.size) {    // then it's one of the special codes:
            FormatHandle.interpretSpecialColorIndex(icvBack.toInt()).toInt()
        } else icvBack.toInt()
    }

    /**
     * return the fg color index for this Area Format
     * @return
     */
    fun geticvFore(): Int {
        return if (icvFore > this.colorTable.size) {    // then it's one of the special codes:
            FormatHandle.interpretSpecialColorIndex(icvFore.toInt()).toInt()
        } else icvFore.toInt()
    }

    fun seticvBack(clr: Int) {
        if (clr > -1 && clr < this.colorTable.size) {
            icvBack = clr.toShort()
            rgbBack = this.colorTable[clr]
            updateRecord()
        } else if (clr == 0x4D) { // special flag, default bg
            icvBack = clr.toShort()
            rgbBack = this.colorTable[0]
            updateRecord()
        }
    }

    /**
     * sets the fill color for the area
     * @param clr color index
     */
    fun seticvFore(clr: Int) {

        // fls= 1 The fill pattern is solid. When solid is specified,
        // rgbFore is the only color rendered, even if rgbBack is specified
        if (clr > -1 && clr < this.colorTable.size) {
            fAuto = false
            fls = 1
            icvFore = clr.toShort()
            rgbFore = this.colorTable[clr]
            // must also set bg to 9
            seticvBack(9)
            //updateRecord();
        } else if (clr == 0x4E) { // default fg
            icvFore = clr.toShort()
            rgbFore = this.colorTable[1]
            updateRecord()
        }
    }

    /**
     * sets the fill color for the area
     * @param clr Hex Color String
     */
    fun seticvFore(clr: String) {
        fAuto = false
        fls = 1
        rgbFore = FormatHandle.HexStringToColor(clr)
        icvFore = FormatHandle.HexStringToColorInt(clr, FormatHandle.colorBACKGROUND).toShort()        // finds best match
        updateRecord()
    }


    /**
     * sets the OOXML settings for this Area Format
     * @param sp
     */
    fun setFromOOXML(sp: SpPr) {
        val f = sp.fill
        if (f != null) {
            this.seticvFore(f.color)
            //this.seticvBack()
            // fls= fill pattern

        }
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -437132913972684937L

        // 20070716 KSC: Need to create new records
        val prototype: XLSRecord?
            get() {
                val af = AreaFormat()
                af.opcode = XLSConstants.AREAFORMAT
                af.data = af.PROTOTYPE_BYTES
                af.init()
                return af
            }                        // changed prototype bytes to ensure has default settings i.e. no pattern and automatic color

        fun getPrototype(type: Int): XLSRecord {
            val af = AreaFormat()
            af.opcode = XLSConstants.AREAFORMAT
            if (type == 0)
                af.data = af.PROTOTYPE_BYTES_1    // default is certain color combo
            else if (type == 1)
                af.data = af.PROTOTYPE_BYTES_2    // ""
            else
                af.data = af.PROTOTYPE_BYTES
            af.init()
            return af
        }
    }

}
