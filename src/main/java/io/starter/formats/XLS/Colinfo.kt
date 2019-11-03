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
package io.starter.formats.XLS

import io.starter.OpenXLS.ExcelTools
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger


/**
 * **Colinfo: Column Formatting Information (7Dh)**<br></br>
 *
 *
 * Colinfo describes the formatting for a column range
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---------------------------------------------------------------
 * 4       colFirst        2       First formatted column (0)
 * 6       colLast         2       Last formatted column (0)
 * 8       colWidth        2       Column width in 1/256 character units
 * 10      ixfe            2       Index to XF for columns
 * 12      grbit           2       Options
 * 14      reserved        1       Must be zero
 *
 *
 * grib options
 * offset	Bits	mask	name		contents
 * ---------------------------------------------------------------
 * 0		0		01h		fHidden		=1 if the column range is hidden
 * 7-1		FEh		UNUSED
 * 1		2-0		07h		iOutLevel	Outline Level of the column
 * 3		08h		Reserved	must be zero
 * 4		10h		iCollapsed	=1 if the col is collapsed in outlining
 * 7-5		E0h		Reserved	must be zero
 *
 * etc.
 *
 * Note: for a discussion of Column widths see:
 * http://support.microsoft.com/?kbid=214123
</pre> *
 */

class Colinfo : XLSRecord(), ColumnRange {
    private var colFirst: Int = 0
    private var colLast: Int = 0
    private var colWidth: Int = 0
    private var grbit: Short = 0
    private var collapsed: Boolean = false
    private var hidden: Boolean = false
    private var outlineLevel = 0

    /**
     * Is this colinfo based on a single column format?
     */
    val isSingleColColinfo: Boolean
        get() = getColFirst() == getColLast()

    /**
     * returns Column Width in Chars or Excel-units
     *
     * @return
     */
    val colWidthInChars: Int
        get() {
            var colwidth = this.getColWidth()
            colwidth = Math.round((colwidth - fudgefactor) / 256).toInt()
            return colwidth
        }

    /**
     * Sets the Ixfe for this record.  For some stupid reason this is the *only* xls record
     * that has it's ixfe in a different place.  Intern time at microsoft I guess.
     */
    override var ixfe: Int
        get() = ixfe
        set(i) {
            this.ixfe = i
            val newxfe = ByteTools.cLongToLEBytes(i)
            val b = this.getData()

            System.arraycopy(newxfe, 0, b!!, 6, 2)
            this.setData(b)
        }

    /**
     * Returns whether the column is collapsed
     *
     * @return
     */
    /**
     * Flag indicating if the outlining of the affected column(s) is in the collapsed state.
     *
     * @param b boolean true if collapsed
     */
    // all previous columns are hidden
    var isCollapsed: Boolean
        get() = collapsed
        set(b) {
            this.collapseIt(b)
            for (i in 0 until this.colFirst) {
                val r = this.sheet!!.getColInfo(i)
                if (r != null && r.outlineLevel == this.getOutlineLevel()) {
                    r.isHidden = b
                }
            }
        }

    /**
     * Returns whether the column is hidden
     *
     * @return
     */
    /**
     * Set whether the column is hidden
     *
     * @param b
     */
    var isHidden: Boolean
        get() = hidden
        set(b) {
            this.hidden = b
            updateGrbit()
            this.workBook!!.usersviewbegin.displayOutlines = true
        }

    override val isSingleCol: Boolean
        get() = this.getColFirst() == this.getColLast()

    /**
     * set last/first cols/rows
     */
    fun setColFirst(c: Int) {
        val b = ByteTools.shortToLEBytes(c.toShort())
        val dt = this.getData()
        System.arraycopy(b, 0, dt!!, 0, 2)
        this.colFirst = c
    }

    override fun getColFirst(): Int {
        return colFirst
    }

    /**
     * set last/first cols/rows
     */
    fun setColLast(c: Int) {
        val b = ByteTools.shortToLEBytes(c.toShort())
        val dt = this.getData()
        System.arraycopy(b, 0, dt!!, 2, 2)
        this.colLast = c
    }

    override fun getColLast(): Int {
        return colLast
    }

    /**
     * Shifts the whole colinfo over the amount of the offset
     */
    fun moveColInfo(offset: Int) {
        this.setColFirst(this.getColFirst() + offset)
        this.setColLast(this.getColLast() + offset)
    }

    override fun init() {
        super.init()
        colFirst = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        colLast = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt()).toInt()
        colWidth = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt()).toInt()
        ixfe = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt()).toInt()
        grbit = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())
        decodeGrbit()
        if (DEBUGLEVEL > 5)
            Logger.logInfo("Col: " + ExcelTools.getAlphaVal(colFirst) + "-" + ExcelTools.getAlphaVal(colLast) + "  ixfe: " + ixfe + " width: " + colWidth)
    }

    /**
     * returns whether a given col
     * is referenced by this Colinfo
     */
    fun inrange(x: Int): Boolean {
        return x <= colLast && x >= colFirst
    }

    fun setColWidthInChars(x: Double) {
        var x = x
        // it's a value that needs to be converted to the appropriate units
        x = (x + fudgefactor).toInt() * 256.0
        val cl = ByteTools.shortToLEBytes(x.toShort())
        System.arraycopy(cl, 0, this.getData()!!, 4, 2)
        colWidth = ByteTools.readShort(this.getData()!![4].toInt(), this.getData()!![5].toInt()).toInt()
        // 20060609 KSC: APPEARS THAT grbit=0 means default column width so must set to either 2 or 6 ---
        // there is NO documentation on this!
        if (grbit.toInt() == 0)
            setGrbit(2)
    }

    /**
     * Set the width of a column or columns in internal units
     * <br></br>Internal units are units of (defaultfontwidth/256)
     * <br></br>Excel column width= default font width/256 * 8.43
     * <br></br>More specifically, column width is in 1/256 of the width of the zero character,
     * using default font (first FONT record in the file)
     *
     * @param int x		new column width
     */
    fun setColWidth(x: Int) {
        val cl = ByteTools.shortToLEBytes(x.toShort())
        System.arraycopy(cl, 0, this.getData()!!, 4, 2)
        colWidth = ByteTools.readUnsignedShort(this.getData()!![4], this.getData()!![5])
        // 20060609 KSC: APPEARS THAT grbit=0 means default column width so must set to either 2 or 6 ---
        // there is NO documentation on this!
        if (grbit.toInt() == 0)
            setGrbit(2)
    }

    /**
     * Get the width of a column in internal units
     * <br></br>Internal units are units of (defaultfontwidth/256)
     * <br></br>Excel column width= default font width/256 * 8.43
     * <br></br>More specifically, column width is in 1/256 of the width of the zero character,
     * using default font (first FONT record in the file)
     */
    fun getColWidth(): Int {
        return colWidth

    }

    /**
     * collapse it is called internally, as you need to call the next column from
     * the colinfo.
     *
     * @param b
     */
    private fun collapseIt(b: Boolean) {
        this.collapsed = b
        updateGrbit()
    }

    /**
     * Set the Outline level (depth) of the column
     *
     * @param x
     */
    fun setOutlineLevel(x: Int) {
        this.outlineLevel = x
        updateGrbit()
        this.sheet!!.guts!!.colGutterSize = 10 + 10 * x
        this.sheet!!.guts!!.maxColLevel = x + 1
    }

    /**
     * This should be run at init()
     * in order to populate grbit values
     */
    private fun decodeGrbit() {
        val grbytes = ByteTools.shortToLEBytes(grbit)
        hidden = grbytes[0] and 0x1 == 0x1
        collapsed = grbytes[1] and 0x10 == 0x10
        outlineLevel = grbytes[1] and 0x7
    }


    fun getGrbit(): Int {
        return grbit.toInt()
    }

    fun setGrbit(grbit: Int) {
        this.grbit = grbit.toShort()
        updateGrbit()
    }

    /**
     * set the grbit to match
     * whatever values have been passed in to
     * modify the grbit functions.  It also updates
     * underlying byte record of the XLSRecord.
     */
    fun updateGrbit() {
        val grbytes = ByteTools.shortToLEBytes(grbit)
        // set whether collapsed or not
        if (collapsed) {
            grbytes[1] = (0x10 or grbytes[1]).toByte()
        } else {
            grbytes[1] = (0xEF and grbytes[1]).toByte()
        }
        // set if hidden
        if (hidden) {
            grbytes[0] = (0x1 or grbytes[0]).toByte()
        } else {
            grbytes[0] = (0xFE and grbytes[0]).toByte()
        }
        // set the outline level
        grbytes[1] = (outlineLevel or grbytes[1]).toByte()
        // reset the grbit and the body rec
        grbit = ByteTools.readShort(grbytes[0].toInt(), grbytes[1].toInt())
        val recdata = this.getData()
        recdata[8] = grbytes[0]
        recdata[9] = grbytes[1]
        this.setData(recdata)
    }

    /**
     * Returns the Outline level (depth) of the column
     *
     * @return
     */
    fun getOutlineLevel(): Int {
        return outlineLevel
    }

    override fun toString(): String {
        return "ColInfo: " + ExcelTools.getAlphaVal(colFirst) + "-" + ExcelTools.getAlphaVal(colLast) + "  ixfe: " + ixfe + " width: " + colWidth
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 3048724897018541459L
        val DEFAULT_COLWIDTH = 2340    // why 2000???? excel reports 2340 ...?


        fun getPrototype(colF: Int, colL: Int, wide: Int, formatIdx: Int): Colinfo {
            val ret = Colinfo()

            ret.originalsize = 12
            // ret.setLabel("COLINFO");
            ret.opcode = XLSConstants.COLINFO
            ret.length = ret.originalsize.toShort()
            val newbytes = ByteArray(ret.originalsize)
            // colF
            var b = ByteTools.shortToLEBytes(colF.toShort())
            newbytes[0] = b[0]
            newbytes[1] = b[1]

            // colL
            b = ByteTools.shortToLEBytes(colL.toShort())
            newbytes[2] = b[0]
            newbytes[3] = b[1]

            // wide
            b = ByteTools.shortToLEBytes(wide.toShort())
            newbytes[4] = b[0]
            newbytes[5] = b[1]

            // XF
            b = ByteTools.shortToLEBytes(formatIdx.toShort())
            newbytes[6] = b[0]
            newbytes[7] = b[1]

            ret.setData(newbytes)
            ret.init()
            return ret
        }

        /**
         * Set the width of a column or columns in Excel units
         *
         * @param double x	new column width in Excel Units
         */
        private val fudgefactor = 0.711    // calc needs fudge factor to equal excel values !!!
    }
}