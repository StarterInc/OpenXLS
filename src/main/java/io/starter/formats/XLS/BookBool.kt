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
 * **BOOKBOOL: Additional Workspace Information (DAh)**<br></br>
 *
 *
 * This record stores information about workspace settings
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       grbit          2      Option flags
 *
 * See book for details of grbit flags, page 425
 *
</pre> *
 */

class BookBool : io.starter.formats.XLS.XLSRecord() {
    internal var grbit: Short = 0

    override fun init() {
        super.init()
        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        if (DEBUGLEVEL > 5)
            Logger.logInfo("BOOKBOOL: " + if (grbit.toInt() == 0) "Save External Links" else "Don't Save External Links")
    }

    companion object {

        private val serialVersionUID = -4544323710670598072L
    }


}