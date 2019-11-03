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
 * **PROTECT: Protection Flag (12h)**<br></br>
 *
 *
 * PROTECT stores the protection state for a sheet or workbook
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       fLock           2       = 1 if the sheet or book is protected
</pre> *
 */
class Protect
/**
 * default constructor
 */
internal constructor() : io.starter.formats.XLS.XLSRecord() {
    private var fLock = 0

    /**
     * returns whether this sheet or workbook is protected
     */
    internal var isLocked: Boolean
        get() = fLock != 0
        set(b) {
            fLock = if (b) 1 else 0
            val data = this.getData()
            data[0] = fLock.toByte()
        }

    init {
        val bs = ByteArray(2)
        opcode = XLSConstants.PROTECT
        length = 2.toShort()
        //   setLabel("PROTECT" + String.valueOf(this.offset));
        setData(bs)
        this.originalsize = 2
    }

    override fun init() {
        super.init()
        fLock = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        if (this.isLocked && DEBUGLEVEL > XLSConstants.DEBUG_LOW)
            Logger.logInfo("Workbook/Sheet Protection Enabled.")
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -2455962145433645632L
    }

}