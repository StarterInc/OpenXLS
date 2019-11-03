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

import io.starter.toolkit.ByteTools


/**
 * **WINDOW1 0x3D: Contains window attributes for a Workbook.**<br></br>
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       xWn         2       Horizontal Position of the window
 * 6       yWn         2       Vertical Position of the window
 * 8       dxWn        2       Width of the window
 * 10      dyWn        2       Height of the window
 * 12      grbit       2       Option Flags
 * 14      itabCur     2       Index of the selected workbook tab (0 based)
 * 16      itabFirst   2       Index of the first displayed workbook tab (0 based)
 * 18      ctabSel     2       Number of workbook tabs that are selected
 * 20      wTabRatio   2       Ratio of the width of the workbook tabs to the width of
 * the horizontal scroll bar; to obtain the ratio, convert to
 * decimal and then divide by 1000
</pre> *
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

class Window1 : io.starter.formats.XLS.XLSRecord() {
    internal var xWn: Short = 0
    internal var yWn: Short = 0
    internal var dxWn: Short = 0
    internal var dyWn: Short = 0
    internal var grbit: Short = 0
    internal var itabCur: Short = 0
    internal var itabFirst: Short = 0
    internal var ctabSel: Short = 0
    internal var wTabRatio: Short = 0
    internal var mybs: Boundsheet? = null

    val currentTab: Int
        get() = itabCur.toInt()

    /**
     * Sets the current tab that is displayed on opening.
     * Note, this is not really the same thing as "selected".
     * The selected parameter is from the Window2 record.  As we
     * don't really have much need to select more than one sheet
     * on output, this method just delselects every other sheet than
     * the one that is passed in, and selects that one in it's Window2
     *
     * @param bs
     */
    fun setCurrentTab(bs: Boundsheet) {
        mybs = bs
        val t = mybs!!.sheetNum
        val bounds = this.workBook!!.workSheets
        for (i in bounds.indices) {
            bounds[i].window2!!.setSelected(false)
        }
        mybs!!.window2!!.setSelected(true)
        val mydata = this.getData()
        itabCur = t.toShort()
        val tabbytes = ByteTools.shortToLEBytes(t.toShort())
        mydata[10] = tabbytes[0]
        mydata[11] = tabbytes[1]
    }

    /**
     * Sets which tab will display furthest to the left in the workbook.  Sheets that have
     * their tabid before this one will be 'pushed off' to the left.  They can be retrieved in the
     * GUI by clicking the left arrow next to the displayed worksheets.
     *
     * @param t
     */
    fun setFirstTab(t: Int) {
        val mydata = this.getData()
        itabFirst = t.toShort()
        val tabbytes = ByteTools.shortToLEBytes(t.toShort())
        mydata[12] = tabbytes[0]
        mydata[13] = tabbytes[1]
    }

    /**
     * Default init method
     */
    override fun init() {
        super.init()
        xWn = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        yWn = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        dxWn = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        dyWn = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        grbit = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())
        itabCur = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt())
        itabFirst = ByteTools.readShort(this.getByteAt(12).toInt(), this.getByteAt(13).toInt())
        ctabSel = ByteTools.readShort(this.getByteAt(14).toInt(), this.getByteAt(15).toInt())
        wTabRatio = ByteTools.readShort(this.getByteAt(16).toInt(), this.getByteAt(17).toInt())

    }

    /**
     * Returns whether the sheet selection tabs should be shown.
     */
    fun showSheetTabs(): Boolean {
        return grbit and 0x20 == 0x20
    }

    /**
     * Sets whether the sheet selection tabs should be shown.
     */
    fun setShowSheetTabs(show: Boolean) {
        if (show)
            grbit = grbit or 0x20
        else
            grbit = grbit and 0x20.inv().toShort()
        val b = ByteTools.shortToLEBytes(grbit)
        this.getData()[8] = b[0]
        this.getData()[9] = b[1]
    }

    companion object {

        /**
         *
         */
        private val serialVersionUID = 2770922305028029883L
    }
}