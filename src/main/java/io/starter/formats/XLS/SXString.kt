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

/**
 * SXString 0x
 * specifies a pivot cache item with a string value.
 *
 *
 * cch (2 bytes): An unsigned integer that specifies the length, in characters, of the XLUnicodeStringNoCch structure in the segment field. If cch is 0xFFFF, segment MUST NOT exist.
 * segment (variable): An XLUnicodeStringNoCch structure that specifies a segment of the string. This exists only if the value of the cch field is different than 0xFFFF.
 */

import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.io.UnsupportedEncodingException
import java.util.Arrays

class SXString : XLSRecord(), XLSConstants, PivotCacheRecord {
    private var cch: Short = 0    // length of segment
    private var segment: String? = null    //specifies a cache item with a string value.

    /**
     * sets the value for the string cache item referenced by this SXString
     * @param s
     */
    // account for encoding bytes + cch
    // 	 (byte[]) [5, 0, 0, 115, 111, 117, 116, 104]	south
    var cacheItem: String?
        get() = segment
        set(s) {
            segment = s
            var strbytes = ByteArray(0)
            if (segment != null) {
                try {
                    strbytes = segment!!.toByteArray(charset(XLSConstants.DEFAULTENCODING))
                } catch (e: UnsupportedEncodingException) {
                    Logger.logInfo("SxString: $e")
                }

            }
            cch = strbytes.size.toShort()
            val nm = ByteTools.shortToLEBytes(cch)
            val data = ByteArray(cch + 3)
            System.arraycopy(nm, 0, data, 0, 2)
            System.arraycopy(strbytes, 0, data, 3, cch.toInt())
            this.setData(data)
        }

    /**
     * return the bytes describing this record, including the header
     * @return
     */
    override val record: ByteArray
        get() {
            val b = ByteArray(4)
            System.arraycopy(ByteTools.shortToLEBytes(this.opcode), 0, b, 0, 2)
            System.arraycopy(ByteTools.shortToLEBytes(this.getData()!!.size.toShort()), 0, b, 2, 2)
            return ByteTools.append(this.getData(), b)

        }

    override fun init() {
        super.init()
        cch = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        if (cch > 0) {
            val encoding = this.getByteAt(2)

            val tmp = this.getBytesAt(3, cch * (encoding + 1))
            try {
                if (encoding.toInt() == 0)
                    segment = String(tmp!!, XLSConstants.DEFAULTENCODING)
                else
                    segment = String(tmp!!, XLSConstants.UNICODEENCODING)
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("SXString.init: $e")
            }

        }
        if (DEBUGLEVEL > 3)
            Logger.logInfo(this.toString())
    }

    override fun toString(): String {
        return "SXString: " + (if (segment != null) segment else "null") +
                Arrays.toString(this.record)
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 9027599480633995587L

        /**
         * create a new minimum SXString
         * @return
         */
        val prototype: XLSRecord?
            get() {
                val sxstring = SXString()
                sxstring.opcode = XLSConstants.SXSTRING
                sxstring.setData(byteArrayOf(-1, -1))
                sxstring.init()
                return sxstring
            }
    }
}
