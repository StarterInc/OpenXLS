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
 * **Pane: Stores the position of window panes. (41h)**
 * <br></br>If the sheet doesn't contain any splits, this record will not occur
 * A sheet can be split in two different ways, with unfrozen or frozen panes.
 * A flag in the WINDOW2 record specifies if the panes are frozen.
 *
 *
 *
 * <pre>
 * offset  size name		contents
 * ---
 * 0 	2	px 			Position of the vertical split (px, 0 = No vertical split):
 * Unfrozen pane: Width of the left pane(s) (in twips = 1/20 of a point)
 * Frozen pane: Number of visible columns in left pane(s)
 * 2 	2 	py			Position of the horizontal split (py, 0 = No horizontal split):
 * Unfrozen pane: Height of the top pane(s) (in twips = 1/20 of a point)
 * Frozen pane: Number of visible rows in top pane(s)
 * 4 	2 	visRow		Index to first visible row in bottom pane(s)
 * 6 	2	visCol 		Index to first visible column in right pane(s)
 * 8 	1 	pActive		Identifier of pane with active cell cursor (see below)
 * [9] 1 				Not used (BIFF5-BIFF8 only, not written in BIFF2-BIFF4)
 *
 * If the panes are frozen, pane 0 is always active, regardless of the cursor position. The correct identifiers for all possible
 * combinations of visible panes are shown in the following pictures.
 * px = 0, py = 0		px = 0, py > 0		px > 0, py = 0		px > 0, py > 0
 * 3				3					3 1					3 1
 * 2										2 0
 *
</pre> *
 */

class Pane : io.starter.formats.XLS.XLSRecord() {
    internal var px: Short = 0
    internal var py: Short = 0
    internal var visRow: Short = 0
    internal var visCol: Short = 0
    internal var pActive: Byte = 0
    internal var bFrozen: Boolean = false
    internal var win2: Window2

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    /**
     * Gets the first visible row of the split pane
     * this is 0 based
     *
     * @return
     */
    val visibleRow: Int
        get() = visRow.toInt()

    /**
     * gets the first visible col of the split pane
     * this is 0 based
     *
     * @return int 0-based first visible column of the split
     */
    val visibleCol: Int
        get() = visCol.toInt()

    /**
     * Returns the position of the row split in twips
     */
    val rowSplitLoc: Int
        get() = py.toInt()


    /**
     * return the address of the TopLeft visible Cell
     *
     * @return
     */
    val topLeftCell: String
        get() = ExcelTools.formatLocation(intArrayOf(visRow.toInt(), visCol.toInt()))

    override fun init() {
        super.init()
        px = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        py = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        visRow = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        visCol = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        pActive = this.getByteAt(8)

    }

    /**
     * sets the 1st visible column height in twips + sets frozen panes off
     *
     * @param col
     * @param nCols
     */
    fun setSplitColumn(col: Int, nCols: Int) {
        setFrozenColumn(col)
        win2.freezePanes = false
        visCol = col.toShort()
        px = nCols.toShort()
        py = 0
        pActive = 3        // TODO: figure this out!
        updateData()
    }

    /**
     * sets split columm + freezes panes
     *
     * @param col
     */
    fun setFrozenColumn(col: Int) {
        win2.freezePanes = true
        visCol = col.toShort()
        px = col.toShort()
        pActive = 0
        updateData()
    }


    /**
     * sets the first visible row + width in twips + sets frozen panes off
     *
     * @param row to start split
     * @param rsz row height in twips
     */
    fun setSplitRow(row: Int, rsz: Int) {
        setFrozenRow(row)
        win2.freezePanes = false
        visRow = row.toShort()
        py = rsz.toShort()
        px = 0
        pActive = 3        // TODO: figure this out!
        updateData()
    }

    /**
     * sets the split row + freezes pane
     *
     * @param row
     */
    fun setFrozenRow(row: Int) {
        win2.freezePanes = true
        visRow = row.toShort()
        py = row.toShort()
        pActive = 0
        updateData()
    }

    protected fun updateData() {
        this.getData()
        var b = ByteTools.shortToLEBytes(px)
        data[0] = b[0]
        data[1] = b[1]
        b = ByteTools.shortToLEBytes(py)
        data[2] = b[0]
        data[3] = b[1]
        b = ByteTools.shortToLEBytes(visRow)
        data[4] = b[0]
        data[5] = b[1]
        b = ByteTools.shortToLEBytes(visCol)
        data[6] = b[0]
        data[7] = b[1]
        data[8] = pActive
    }

    fun setWindow2(win2: Window2) {
        this.win2 = win2
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 5314818835334217157L

        val prototype: XLSRecord?
            get() {
                val p = Pane()
                p.opcode = XLSConstants.PANE
                p.setData(p.PROTOTYPE_BYTES)
                p.init()
                return p
            }
    }
}