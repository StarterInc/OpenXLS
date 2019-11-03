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


/**
 * **WINDOW2 0x23E: Contains window attributes for a Sheet.**<br></br>
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       grbit       2       Option Flags
 * 6       rwTop       2       Top row visible in the window
 * 8       colLeft     2       Leftmost visible col in window
 * 10      icvHdr      4       Index to color val for row/col headings & grids
 * 14      wScaleSLV   2       Zoom mag in page break preview
 * 16      wScaleNorm  2       Zoom mag in Normal preview
 * 18      reserved    4
 *
 * grbit Option flags
 * 0 0001H 0 = Show formula results 1 = Show formulas
 * 1 0002H 0 = Do not show grid lines 1 = Show grid lines
 * 2 0004H 0 = Do not show sheet headers 1 = Show sheet headers
 * 3 0008H 0 = Panes are not frozen 1 = Panes are frozen (freeze)
 * 4 0010H 0 = Show zero values as empty cells 1 = Show zero values
 * 5 0020H 0 = Manual grid line colour 1 = Automatic grid line colour
 * 6 0040H 0 = Columns from left to right 1 = Columns from right to left
 * 7 0080H 0 = Do not show outline symbols 1 = Show outline symbols
 * 8 0100H 0 = Keep splits if pane freeze is removed 1 = Remove splits if pane freeze is removed
 * 9 0200H 0 = Sheet not selected 1 = Sheet selected (BIFF5-BIFF8)
 * 10 0400H 0 = Sheet not visible 1 = Sheet visible (BIFF5-BIFF8)
 * 11 0800H 0 = Show in normal view 1 = Show in page break preview (BIFF8)    </pre>
 *
 * @see WorkBook
 *
 * @see BOUNDSHEET
 *
 * @see INDEX
 *
 * @see DBCELL
 *
 * @see ROW
 *
 * @see Cell
 *
 * @see XLSRecord
 */

class Window2 : io.starter.formats.XLS.XLSRecord() {
    internal var grbit = -1
    internal var rwTop = -1
    internal var colLeft = -1
    internal var icvHdr = -1
    internal var wScaleSLV = -1
    internal var wScaleNorm = -1

    /**
     * returns the String Address of first visible cell on the sheet
     *
     * @return String
     */
    val topLeftCell: String
        get() = ExcelTools.formatLocation(intArrayOf(rwTop, colLeft))

    var showFormulaResults: Boolean
        get() = grbit and BITMASK_SHOWFORMULARESULTS == BITMASK_SHOWFORMULARESULTS
        set(b) {
            if (b)
                grbit = grbit or BITMASK_SHOWFORMULARESULTS
            else
                grbit = grbit and BITMASK_SHOWFORMULARESULTS.inv()
            this.setGrbit()
        }

    var showGridlines: Boolean
        get() = grbit and BITMASK_SHOWGRIDLINES == BITMASK_SHOWGRIDLINES
        set(b) {
            if (b)
                grbit = grbit or BITMASK_SHOWGRIDLINES
            else
                grbit = grbit and BITMASK_SHOWGRIDLINES.inv()
            this.setGrbit()
        }

    var showSheetHeaders: Boolean
        get() = grbit and BITMASK_SHOWSHEETHEADERS == BITMASK_SHOWSHEETHEADERS
        set(b) {
            if (b)
                grbit = grbit or BITMASK_SHOWSHEETHEADERS
            else
                grbit = grbit and BITMASK_SHOWSHEETHEADERS.inv()
            this.setGrbit()
        }


    //return wScaleSLV;
    // wScaleSLV 10,11
    // wScaleNorm 12,13;
    var scaleNorm: Int
        get() = this.wScaleNorm
        set(zm) {
            wScaleNorm = zm
            val data = this.getData()
            val b = ByteTools.shortToLEBytes(zm.toShort())
            System.arraycopy(b, 0, data!!, 12, 2)
            this.setData(data)
        }


    var showZeroValues: Boolean
        get() = grbit and BITMASK_SHOWZEROVALUES == BITMASK_SHOWZEROVALUES
        set(b) {
            if (b)
                grbit = grbit or BITMASK_SHOWZEROVALUES
            else
                grbit = grbit and BITMASK_SHOWZEROVALUES.inv()
            this.setGrbit()
        }

    var showOutlineSymbols: Boolean
        get() = grbit and BITMASK_SHOWOUTLINESYMBOLS == BITMASK_SHOWOUTLINESYMBOLS
        set(b) {
            if (b)
                grbit = grbit or BITMASK_SHOWOUTLINESYMBOLS
            else
                grbit = grbit and BITMASK_SHOWOUTLINESYMBOLS.inv()
            this.setGrbit()
        }

    /**
     * true if sheet is in normal view mode (false if in page break preview mode)
     *
     * @return
     */
    var showInNormalView: Boolean
        get() = grbit and BITMASK_SHOWINPRINTPREVIEW == 0
        set(b) {
            if (b)
                grbit = grbit and BITMASK_SHOWINPRINTPREVIEW.inv()
            else
                grbit = grbit or BITMASK_SHOWINPRINTPREVIEW
            this.setGrbit()
        }

    var freezePanes: Boolean
        get() = grbit and BITMASK_FREEZEPANES == BITMASK_FREEZEPANES
        set(b) {
            if (b)
                grbit = grbit or BITMASK_FREEZEPANES
            else
                grbit = grbit and BITMASK_FREEZEPANES.inv()
            this.setGrbit()
        }

    var manualGridLineColor: Boolean
        get() = grbit and BITMASK_GRIDLINECOLOR == BITMASK_GRIDLINECOLOR
        set(b) {
            if (b)
                grbit = grbit or BITMASK_GRIDLINECOLOR
            else
                grbit = grbit and BITMASK_GRIDLINECOLOR.inv()
            this.setGrbit()
        }

    override fun init() {
        super.init()
        val s1: Short
        val s2: Short

        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        rwTop = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt()).toInt()
        colLeft = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt()).toInt()

        s1 = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        s2 = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())
        icvHdr = ByteTools.readInt(s1.toInt(), s2.toInt())

        // the following do not necessarily exist in VB-manipulated windows
        if (this.length > 10) {
            wScaleSLV = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt()).toInt()
            wScaleNorm = ByteTools.readShort(this.getByteAt(12).toInt(), this.getByteAt(13).toInt()).toInt()
        }
    }

    internal fun setSelected(b: Boolean) {
        if (b)
            this.getData()[1] = this.getData()[1] or 0x2
        else
            this.getData()[1] = this.getData()[1] and 0xFD
        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
    }

    // add get/set for Window2 options
    fun setGrbit() {
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(grbit.toShort())
        System.arraycopy(b, 0, data!!, 0, 2)
        this.setData(data)

    }

    companion object {
        /**
         *
         */
        private val serialVersionUID = -8316509425117672619L

        //20060308 KSC: Added for get/set access to Window2 options
        internal val BITMASK_SHOWFORMULARESULTS = 0x0001
        internal val BITMASK_SHOWGRIDLINES = 0x0002
        internal val BITMASK_SHOWSHEETHEADERS = 0x0004
        internal val BITMASK_FREEZEPANES = 0x0008
        internal val BITMASK_SHOWZEROVALUES = 0x0010
        internal val BITMASK_GRIDLINECOLOR = 0x0020
        internal val BITMASK_COLUMNDIRECTION = 0x0040
        internal val BITMASK_SHOWOUTLINESYMBOLS = 0x0080
        internal val BITMASK_KEEPSPLITS = 0x0100
        internal val BITMASK_SHEETSELECTED = 0x0200
        internal val BITMASK_SHEETVISIBLE = 0x0400
        internal val BITMASK_SHOWINPRINTPREVIEW = 0x0800
    }
    /*	// TODO: finish these    
    static final int BITMASK_COLUMNDIRECTION= 0x0040;
    static final int BITMASK_KEEPSPLITS= 0x0100;
    static final int BITMASK_SHEETSELECTED= 0x0200;
    static final int BITMASK_SHEETVISIBLE= 0x0400;
*/

}