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


/**
 * **RRTabID:  Revision Tab ID Record**<br></br>
 *
 *
 * The RRTabId record specifies an array of unique sheet identifiers,
 * each of which is associated with a sheet in the workbook.
 * The order of the sheet identifiers in the array matches the order of
 * the BoundSheet8 records as they appear in the Globals substream.
 */
class TabID : io.starter.formats.XLS.XLSRecord() {
    /**
     * @return Returns the tabIDs.
     */
    var tabIDs = CompatibleVector()
        internal set

    /**
     * Default init
     */
    override fun init() {
        super.init()
        var i = 0
        while (i < this.length - 4) {
            val s = ByteTools.readShort(this.getByteAt(i).toInt(), this.getByteAt(i + 1).toInt())
            val sh = java.lang.Short.valueOf(s)
            tabIDs.add(sh)
            i += 2
        }
    }


    /**
     * Looks sequentally at the tabIDs and
     * makes a new one larger than the previous largest...
     */
    internal fun removeRecord() {
        var largest: Short = 0
        for (i in tabIDs.indices) {
            val sh = tabIDs[i] as Short
            if (sh > largest) largest = sh
        }
        tabIDs.remove(java.lang.Short.valueOf(largest))
        this.updateRecord()
    }


    /**
     * Looks sequentally at the tabIDs and
     * makes a new one larger than the previous largest...
     */
    internal fun addNewRecord() {
        var largest: Short = 0
        for (i in tabIDs.indices) {
            val sh = tabIDs[i] as Short
            if (sh > largest) largest = sh
        }
        largest += 0x1
        val sh = java.lang.Short.valueOf(largest)
        tabIDs.add(sh)
        this.updateRecord()
    }


    /**
     * This DOES NOT do what was expected.  Sheet order is soley based off of physical Boundsheet ordering
     * in the output file.  I'm keeping this code in here in case we start supporting revisions.
     */
    private fun changeOrder(sheet: Int, newpos: Int): Boolean {
        val sz = tabIDs.size
        if (sheet < 0 || newpos < 0 || sheet >= sz || newpos >= sz) {
            Logger.logWarn("changing Sheet order failed: invalid Sheet Index: $sheet:$newpos")
            return false
        }
        val b = tabIDs[sheet]
        tabIDs.remove(b)
        tabIDs.insertElementAt(b, newpos)
        this.updateRecord()
        return true
    }

    /**
     * Updates the underlying byte array with the ordered tabId's
     * Call after any modification to this record
     */
    fun updateRecord() {
        val newlen = (tabIDs.size * 2).toShort()
        val newbody = ByteArray(newlen)
        var counter = 0
        for (i in tabIDs.indices) {
            val sh = tabIDs[i] as Short
            val b = ByteTools.shortToLEBytes(sh)
            newbody[counter] = b[0]
            newbody[counter + 1] = b[1]
            counter += 2
        }
        this.setData(newbody)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 722748113519841817L
    }
}