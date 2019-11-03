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


/**
 * **USERSVIEWEND: End Custom View Settings (1ABh)**<br></br>
 *
 *
 * USERSVIEWEND marks the end of a custom view for the sheet
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       fValid          2       = 1 if the settings saved are valid
 *
</pre> *
 */
class Usersviewend : io.starter.formats.XLS.XLSRecord() {
    internal var fValid = -1

    /**
     * return whether the settings
     * for this user view are valid
     */
    val isValid: Boolean
        get() = fValid == 1

    override fun init() {
        super.init()
        fValid = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -3120417369791123931L
    }

}