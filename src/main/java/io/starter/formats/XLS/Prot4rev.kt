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
 * **PROT4REV: Protection Flag (12h)**<br></br>
 *
 *
 * PROT4REV stores the protection state for a shared workbook
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       fRevLock        2       = 1 if the Sharing with
 * track changes option is on
</pre> *
 */
class Prot4rev : io.starter.formats.XLS.XLSRecord() {
    internal var fRevLock = -1

    /**
     * returns whether this sheet or workbook is protected
     */
    internal val isLocked: Boolean
        get() = fRevLock > 0

    override fun init() {
        super.init()
        fRevLock = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        if (this.isLocked) {
            Logger.logInfo("Shared Workbook Protection Enabled.")
            // throw new InvalidFileException("Shared Workbook Protection Enabled.  Unsupported file format.");
        }
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 681662633243537043L
    }

}