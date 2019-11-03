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

import io.starter.toolkit.Logger

/**
 * SXFORMAT FBh: This record contains formatting data
 *
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       rlType      2       0x0 = clear
 * 0x1 = format applied
 * 6       cbData      2       length of data
</pre> *
 */


class Sxformat : XLSRecord(), XLSConstants {
    internal var data: ByteArray? = null

    override fun init() {
        super.init()
        if (this.length <= 0) {  // Is this record populated?
            if (DEBUGLEVEL > -1) Logger.logInfo("no data in Sxformat")
        } else { // parse out all the fields
            // Logger.logInfo(ExcelTools.getRecordByteDef(this));
        }
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = -8702313274711819140L
    }


}