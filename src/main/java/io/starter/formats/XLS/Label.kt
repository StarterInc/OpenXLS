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
import io.starter.toolkit.Logger


/**
 * **Label: BiffRec Value, String Constant (204h)**<br></br>
 * The Label record describes a cell that contains a string.
 * The String length must be in the range 000h-00ffh (0-255).
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2       Row Number
 * 6       col         2       Column Number of the RK record
 * 8       ixfe        2       Index to XF cell format record
 * 10      cch         2       Length of the string
 * 12      rgch        var     The String
</pre> *
 *
 * @see LABELSST
 *
 * @see STRING
 *
 * @see RSTRING
 */

class Label : XLSCellRecord() {
    internal var cch: Int = 0
    internal var `val`: String? = null

    override// byte[] newlenbytes = bto
    var stringVal: String?
        get() = `val`
        set(v) {
            `val` = v
            val newstrlen = v.length
            val newbytes = ByteArray(newstrlen + 8)
            System.arraycopy(getData()!!, 0, newbytes, 0, 6)
            val blen = ByteTools.cLongToLEBytes(newstrlen)
            System.arraycopy(blen, 0, newbytes, 6, 2)
            val strbytes = v.toByteArray()
            System.arraycopy(strbytes, 0, newbytes, 8, newstrlen)
            this.setData(newbytes)
            this.init()
        }

    override fun init() {
        super.init()
        var s: Short
        val s1: Short
        // get the row, col and ixfe information
        s = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        rw = s.toInt()
        s = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        colNumber = s
        s = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        ixfe = s.toInt()
        // get the length of the string
        s1 = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        cch = s1.toInt()
        if (this.getByteAt(8) > 1) { // TODO KSC: Is this the correct indicator to read bytes as unicode??
            val namebytes = this.getBytesAt(8, this.length - 8)
            `val` = String(namebytes!!)
        } else {
            // 20060809 KSC: read correct bytes to interpret as unicode
            try {
                var thistr: Unicodestring? = null
                val tmpBytes = this.getBytesAt(6, cch * 2 + 4)  // i.e. (cch * 2) - 2
                thistr = Unicodestring()
                thistr.init(tmpBytes!!, false)
                `val` = thistr.toString()
            } catch (e: Exception) {
                Logger.logWarn("ERROR Label.init: decoding string failed: $e")
            }

        }
        this.isValueForCell = true
        this.isString = true
    }

    internal fun setStringVal(v: String, b: Boolean) {
        this.stringVal = v
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -2921430854162954640L
    }
}