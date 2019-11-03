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


class SXBool : XLSRecord(), XLSConstants, PivotCacheRecord {
    internal var bool: Boolean = false

    /**
     * return the bytes describing this record, including the header
     *
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
        if (DEBUGLEVEL > 3) Logger.logInfo("SXBool -")
        bool = this.getByteAt(0).toInt() == 0x1
    }

    override fun toString(): String {
        return "SXBool: $bool"
    }

    fun setBool(b: Boolean) {
        bool = b
        if (bool)
            this.getData()[0] = 0x1
        else
            this.getData()[0] = 0
    }

    fun getBool(): Boolean {
        return bool
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 9027599480633995587L

        /**
         * create a new SXBool record
         *
         * @return
         */
        val prototype: XLSRecord?
            get() {
                val sxbool = SXBool()
                sxbool.opcode = XLSConstants.SXBOOL
                sxbool.setData(ByteArray(2))
                return sxbool
            }
    }
}
