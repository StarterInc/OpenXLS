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

import java.nio.charset.StandardCharsets
import java.util.ArrayList

/**
 * **CRN (005Ah)**<br></br>
 *
 *
 * This record stores the contents of an external cell or cell range. An external cell range has one row only. If a cell range
 * spans over more than one row, several CRN records will be created.
 *
 *
 * <pre>
 * offset  size    contents
 * ---
 * 0 		1 		Index to last column inside of the referenced sheet (lc)
 * 1 		1 		Index to first column inside of the referenced sheet (fc)
 * 2 		2 		Index to row inside of the referenced sheet
 * 4 		var.	List of lc-fc+1 cached values
</pre> *
 */

class Crn : XLSRecord() {
    private var lc: Byte = 0
    private var fc: Byte = 0
    private var rowIndex: Int = 0
    private val cachedValues = ArrayList()

    override fun init() {
        super.init()
        lc = this.getByteAt(0)
        fc = this.getByteAt(1)
        rowIndex = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt()).toInt()
        var pos = 4
        for (i in 0 until lc - fc + 1) {
            try {
                val type = this.getByteAt(pos++).toInt()
                when (type) {
                    0    // empty
                    -> pos += 8
                    1    // numeric
                    -> {
                        cachedValues.add(ByteTools.eightBytetoLEDouble(this.getBytesAt(pos, 8)!!))
                        pos += 8
                    }
                    2 // string
                    -> {
                        val ln = ByteTools.readShort(this.getByteAt(pos).toInt(), this.getByteAt(pos + 1).toInt())
                        val encoding = this.getByteAt(pos + 2)
                        pos += 3
                        if (encoding.toInt() == 0) {
                            cachedValues.add(String(this.getBytesAt(pos, ln.toInt())!!))
                            pos += ln.toInt()
                        } else {// unicode
                            cachedValues.add(String(this.getBytesAt(pos, ln * 2)!!, StandardCharsets.UTF_16LE))
                            pos += ln * 2
                        }
                    }
                    4 // boolean
                    -> {
                        cachedValues.add(java.lang.Boolean.valueOf(this.getByteAt(pos + 1).toInt() == 1))
                        pos += 8
                    }
                    16 // error
                    -> {
                        cachedValues.add("Error Code: " + this.getByteAt(pos + 1))
                        pos += 8
                    }
                }
            } catch (e: Exception) {

            }

        }
    }

    override fun toString(): String {
        var ret = "CRN: lc=$lc fc=$fc rowIndex=$rowIndex"
        for (i in cachedValues.indices) {
            ret += " (" + i + ")=" + cachedValues.get(i)
        }
        return ret
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 3162130963170092322L
    }
}
