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
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger

import java.io.ByteArrayOutputStream
import java.io.IOException
import java.util.ArrayList


/**
 * **Mulrk: Multiple Rk Cells (BDh)**<br></br>
 * This record stores up to 256 Rk equivalents in
 *
 *
 * TODO: check compatibility with Excel2007 MAXCOLS
 *
 *
 * a space-saving format.
 * offset  name        size    contents
 * ---
 * 4       rw          2       Row Number
 * 6       colFirst    2       Column Number of the first col of multiple Rk record
 * 8       rgrkrec     var     Array of 6-byte RkREC objects
 * var     colLast     2       Last Column containing the RkREC object
 *
 *
 * @see Rk
 */

class Mulrk internal constructor() : io.starter.formats.XLS.XLSRecord(), Mul {
    internal var removed = false

    internal var colFirst: Short = 0
    internal var colLast: Int = 0
    internal var datalen: Int = 0
    internal var numRkRecs = 0
    internal var rkrecs: MutableList<*>

    override val recs: List<*>
        get() = rkrecs

    /**
     * whether this mul was removed from the SheetRecs already
     *
     * @return
     */
    override fun removed(): Boolean {
        return removed
    }

    /**
     * populate the MULRk with its data, as well as creating
     * multiple Rk records per the Rk array.
     */
    override fun init() {
        super.init()
        var datalen = this.getData()!!.size    //getLength();

        if (datalen <= 0) {
            if (DEBUGLEVEL > -1) Logger.logInfo("no data in MULRk")
        } else {
            super.initRowCol()
            var s = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
            colFirst = s
            colNumber = colFirst
            s = ByteTools.readShort(this.getByteAt(datalen - 2).toInt(), this.getByteAt(datalen - 1).toInt())
            colLast = s.toInt()
            // get the records data only
            datalen = datalen - 6
            val rkdatax = this.getBytesAt(4, datalen)
            numRkRecs = datalen / 6
            //rkrecs = new Rk[numRkRecs]; Now its a vector
            rkrecs = ArrayList(numRkRecs)
            var reccount = 0
            var rkcol = colNumber.toInt()
            // iterate through the rk data array and create
            // a new 6-byte Rk for each.
            var i = 4
            while (i < rkdatax!!.size) {
                val rkd = this.getBytesAt(i, 6)
                val r = Rk()
                r.init(rkd, rw, rkcol++)
                if (DEBUGLEVEL > 5) Logger.logInfo(" rk@" + (rkcol - 1) + ":" + r.stringVal)
                i += 6
                if (reccount == numRkRecs) break
                r.myMul = this
                r.setSheet(sheet)
                r.streamer = streamer
                rkrecs.add(r)
                reccount++
            }
            if (DEBUGLEVEL > 5) Logger.logInfo("Done adding Rk recs to: " + this.cellAddress)
        }
    }

    internal fun deleteRk(rik: Rk) {
        rkrecs.remove(rik)
    }

    internal fun addRk(rik: Rk) {
        rkrecs.add(rik)
    }

    internal fun getColFirst(): Int {
        return colFirst.toInt()
    }

    /**
     * get a handle to a specific Rk for use in updating values
     */
    internal fun getRk(rnum: Int): Rk {
        return rkrecs[rnum] as Rk
    }

    /*
       Changes the range of
    **/
    fun splitMulrk(splitcol: Int): Mulrk? {
        if (splitcol < colFirst || splitcol > colLast) return null
        val newmul = Mulrk()
        newmul.colFirst = splitcol.toShort()
        newmul.colLast = this.colLast
        this.colLast = splitcol - 1
        val rkr = this.recs.iterator()
        while (rkr.hasNext()) {
            val r = rkr.next() as Rk
            if (r.rowNumber >= splitcol) {
                this.deleteRk(r)
                newmul.addRk(r)
            }
        }
        newmul.opcode = opcode
        newmul.length = length
        return newmul
    }


    /**
     * Remove an Rk from the record along with all of the
     * folloing Rks.  Returns a CompatableVector
     * of RKs that have been cut off from the Mulrk.
     * this is kinda deprecated because of the splitMulrk(),
     * but could prove to be useful later...
     */
    internal fun removeRk(rok: Rk): CompatibleVector {
        val rez = CompatibleVector()
        // set the new last col of the Mulrk
        colLast = rok.colNumber - 1
        rkrecs.remove(rok)
        val z = rkrecs.size - 1
        for (i in z downTo 0) {
            val rec = rkrecs[i] as Rk
            if (rec.colNumber > colLast) {
                rez.add(rec)
                rkrecs.removeAt(i)
            }
        }
        this.updateRks()
        return rez
    }

    /**
     * set the row
     */
    fun setRow(i: Int) {
        val r = ByteTools.shortToLEBytes(i.toShort())
        System.arraycopy(r, 0, this.getData()!!, 0, 2)
        rw = i
    }

    /**
     * Update the underlying byte array for the MULRk record
     * after changes have been made to individual Rk records.
     */
    internal fun updateRks() {
        if (this.recs.size < 1) {
            this.sheet!!.removeRecFromVec(this)
            return
        }
        val tmp = ByteArray(4)
        System.arraycopy(getData()!!, 0, tmp, 0, 4)
        val it = this.recs.iterator()
        val out = ByteArrayOutputStream()
        try {
            out.write(tmp)
            // loop through the Rks and copy their bytes to the MULRk byte array.
            while (it.hasNext()) {
                out.write((it.next() as Rk).bytes!!)
            }
            out.write(ByteTools.shortToLEBytes(colLast.toShort()))
        } catch (a: IOException) {
            Logger.logInfo("parsing record continues failed: $a")
        }

        this.setData(out.toByteArray())
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 1438740082267768419L
    }

}