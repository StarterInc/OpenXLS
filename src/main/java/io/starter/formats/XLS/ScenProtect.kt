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
 * **SCENPROTECT: Protection Flag (DDh)**<br></br>
 *
 *
 * SCENPROTECT stores the protection state for a sheet or workbook
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       fLock           2       = 1 if the scenarios are protected
</pre> *
 */

class ScenProtect
/**
 * default constructor
 */
internal constructor() : io.starter.formats.XLS.XLSRecord() {
    internal var fLock = -1

    /**
     * returns whether this sheet or workbook is protected
     */
    internal var isLocked: Boolean
        get() = fLock > 0
        set(b) {
            val data = this.getData()
            if (b)
                data[0] = 1
            else
                data[0] = 0
        }

    init {
        val bs = ByteArray(2)
        bs[0] = 1
        bs[1] = 0
        opcode = XLSConstants.SCENPROTECT
        length = 2.toShort()
        //  setLabel("SCENPROTECT" + String.valueOf(this.offset));
        this.setData(bs)
        this.originalsize = 2
    }

    override fun init() {
        super.init()
        fLock = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        if (this.isLocked && DEBUGLEVEL > 3) Logger.logInfo("Scenario Protection Enabled.")
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -3722344748446193860L
    }
}