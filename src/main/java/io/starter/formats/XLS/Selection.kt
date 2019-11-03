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

import java.io.Serializable


/**
 * **SELECTION 0x1Dh: Describes the currently selected area of a Sheet.**<br></br>
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       pnn         1       Number of the pane described
 * 5       rwAct       2       Row of the active Cell
 * 7       colAct      2       Col of the active Cell
 * 9       irefAct     2       Ref number of the active Cell
 * 11      cref        2       Number of refs in the selection
 * 13      rgref       var     Array of refs
 *
</pre> *
 *
 * @see WorkBook
 *
 * @see ROW
 *
 * @see Cell
 *
 * @see XLSRecord
 */
class Selection : io.starter.formats.XLS.XLSRecord() {
    internal var pnn: Short = 0
    internal var rwAct: Short = 0
    internal var colAct: Short = 0
    internal var irefAct: Short = 0
    internal var cref: Short = 0
    internal var refs: Array<rgref>

    override fun init() {
        super.init()
        rwAct = ByteTools.readShort(this.getByteAt(1).toInt(), this.getByteAt(2).toInt())
        colAct = ByteTools.readShort(this.getByteAt(3).toInt(), this.getByteAt(4).toInt())
        irefAct = ByteTools.readShort(this.getByteAt(5).toInt(), this.getByteAt(6).toInt())
        cref = ByteTools.readShort(this.getByteAt(7).toInt(), this.getByteAt(8).toInt())


        // cref is count of ref structs -- each one is 6 bytes
        refs = arrayOfNulls<rgref>(cref)
        var ctr = 9
        for (i in 0 until cref.toInt()) {
            val b1 = ByteArray(6)
            for (x in 0..5) {
                b1[x] = this.getByteAt(ctr++)
            }
            refs[i] = rgref(b1)
        }
        val checklen = cref * 6 + 9
        // Logger.logInfo("Done adding " + String.valueOf(cref) + " rgrefs to SELECTION.  Record should be: " + String.valueOf(checklen) + " bytes long.");
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 2949920585425685061L
    }

}


/**
 * Ref Structures -- allows handling of multiple
 * selection information
 */
internal class rgref(var info: ByteArray) : Serializable {
    companion object {
        /**
         * serialVersionUID
         */
        private const val serialVersionUID = 7261215340827889609L
    }
}