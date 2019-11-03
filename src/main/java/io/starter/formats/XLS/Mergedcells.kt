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


import io.starter.OpenXLS.CellRange
import io.starter.OpenXLS.WorkBookHandle
import io.starter.OpenXLS.WorkSheetHandle
import io.starter.toolkit.ByteTools
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger


/**
 * **Mergedcells: Merged Cells for Sheet (E5h)**<br></br>
 *
 *
 * Merged Cells record
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 0	 					var		BiffRec range addresses
 *
</pre> *
 */
class Mergedcells : io.starter.formats.XLS.XLSRecord() {
    private var nummerges = 0
    private var ranges = CompatibleVector()

    /**
     * returns the Merged BiffRec Ranges for this WorkSheet
     *
     * @return an array of CellRanges each containing Merged Cells.
     */
    val mergedRanges: Array<CellRange>?
        get() {
            if (ranges.size < 1) return null
            val ret = arrayOfNulls<CellRange>(ranges.size)
            ranges.toTypedArray()
            return ret
        }

    override fun init() {
        super.init()
        if (DEBUGLEVEL > 5) Logger.logInfo("Mergedcells record.")
    }

    /**
     * Un-merge a CellRange
     *
     * @param rng
     */
    fun removeCellRange(rng: CellRange) {
        this.ranges.remove(rng)
        this.update()
    }

    /**
     * merge a CellRange
     *
     * @param rng
     */
    fun addCellRange(rng: CellRange) {
        //  rng.setIsmerge(true);
        this.ranges.add(rng)
    }

    /**
     * update the underlying bytes with any new CellRange info
     */
    fun update() {
        if (ranges.size > MAXRANGES) {
            this.handleMultiRec()
        }
        this.nummerges = ranges.size
        var datasz = nummerges * 8
        datasz += 2
        data = ByteArray(datasz)
        // get the number of CellRanges
        val szbt = ByteTools.shortToLEBytes(ranges.size.toShort())
        data[0] = szbt[0]
        data[1] = szbt[1]
        var pos = 2
        if (DEBUGLEVEL > XLSConstants.DEBUG_LOW) Logger.logInfo("updating Mergedcell with $nummerges merges.")
        for (t in ranges.indices) {
            val thisrng = ranges[t] as CellRange
            // why call this now??? much overhead ...			if(!thisrng.update())
            // note: range is updated upon setRange ...
            //				return;
            val rints = thisrng.rowInts
            val cints = thisrng.colInts

            //Logger.logInfo(rints[0]+",");
            //for(int x=0;x<cints.length;x++)Logger.logInfo(cints[x]+",");
            //Logger.logInfo();

            val rowmin = ByteTools.shortToLEBytes((rints[0] - 1).toShort())
            data[pos++] = rowmin[0]
            data[pos++] = rowmin[1]

            val rowmax = ByteTools.shortToLEBytes((rints[rints.size - 1] - 1).toShort())
            data[pos++] = rowmax[0]
            data[pos++] = rowmax[1]

            val colmin = ByteTools.shortToLEBytes(cints[0].toShort())
            data[pos++] = colmin[0]
            data[pos++] = colmin[1]

            val colmax = ByteTools.shortToLEBytes(cints[cints.size - 1].toShort())
            data[pos++] = colmax[0]
            data[pos++] = colmax[1]

        }
        this.setData(data)
    }

    /**
     * Mergedranges do not use Continues but instead
     * there are multiple Mergedranges.
     *
     * @return
     */
    internal fun handleMultiRec() {
        //if(true)return;
        if (ranges.size < MAXRANGES) return
        this.nummerges = MAXRANGES
        val mcfresh = Mergedcells.prototype as Mergedcells?
        val substa = ranges.subList(MAXRANGES, ranges.size)
        val ita = substa.iterator()
        while (ita.hasNext()) {
            mcfresh!!.addCellRange(ita.next() as CellRange)
        }
        val removes = mcfresh!!.ranges.iterator()
        while (removes.hasNext()) ranges.remove(removes.next())
        this.workBook!!.addRecord(mcfresh, false)

        val idx = this.recordIndex + 1
        mcfresh.setSheet(this.sheet)
        this.sheet!!.addMergedCellsRec(mcfresh)
        //Logger.logInfo("ADDING Mergedcells at idx: " + idx);
        this.streamer!!.addRecordAt(mcfresh, idx)
        mcfresh.init()
        mcfresh.update()
    }

    /**
     * Initialize the CellRanges containing the Merged Cells.
     *
     * @param the workbook containing the cells
     */
    fun initCells(wbook: WorkBookHandle) {
        nummerges = ByteTools.readShort(this.getByteAt(0).toInt(), getByteAt(1).toInt()).toInt()
        ranges = CompatibleVector()
        var pos = 2 // pointer to the indexes
        for (x in 0 until nummerges) {
            val cellcoords = IntArray(4)
            cellcoords[0] = ByteTools.readShort(getByteAt(pos++).toInt(), getByteAt(pos++).toInt()).toInt() // first col onebased
            cellcoords[2] = ByteTools.readShort(getByteAt(pos++).toInt(), getByteAt(pos++).toInt()).toInt() // last col onebased
            cellcoords[1] = ByteTools.readShort(getByteAt(pos++).toInt(), getByteAt(pos++).toInt()).toInt() // first row
            cellcoords[3] = ByteTools.readShort(getByteAt(pos++).toInt(), getByteAt(pos++).toInt()).toInt() // last row
            try {
                // for(int xd=0;xd<cellcoords.length;xd++)
                //    Logger.logInfo("" + cellcoords[xd]+",");
                val shtr = wbook.getWorkSheet(this.sheet!!.sheetName)

                // TODO: testing -- this saves about 30MB in parsing the Reflexis 700+ sheet problem
                val cr = CellRange(shtr!!, cellcoords, false)

                //	Logger.logInfo(" init: " + cr);;
                //	Logger.logInfo(x);
                cr.workBook = wbook
                ranges.add(cr)
                val ch = cr.cellRecs
                var aMul: Mulblank? = null
                for (t in ch.indices) {
                    // set the range of merged cells
                    if (ch[t] != null) {
                        if (ch[t].opcode == XLSConstants.MULBLANK) {
                            if (aMul == ch[t])
                                continue    // skip- already handled
                            else {
                                aMul = ch[t] as Mulblank
                            }
                        }
                        ch[t].mergeRange = cr
                    }
                }
            } catch (e: Throwable) {
                //	Logger.logWarn("initializing Merged Cells failed: " + e);
            }

        }
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 6638569392267433468L
        var MAXRANGES = 1024

        // for larger records, we need to read in from a file...
        val prototype: XLSRecord?
            get() {
                val newrec = Mergedcells()
                newrec.opcode = XLSConstants.MERGEDCELLS
                newrec.setData(byteArrayOf(0x0, 0x0, 0x0, 0x0))
                return newrec
            }
    }


}