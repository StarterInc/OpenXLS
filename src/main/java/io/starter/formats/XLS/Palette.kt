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
import io.starter.toolkit.CompatibleVector


/**
 * **Palette: Defined Colors (92h)**<br></br>
 *
 *
 * Describe Colors in the file.
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       ccv             2       count of  colors
 * var     rgch            var     4 byte color data
 *
</pre> *
 */
class Palette : io.starter.formats.XLS.XLSRecord() {
    internal var colorvect = CompatibleVector()
    internal var ccv = -1

    override fun init() {
        super.init()
        ccv = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt()).toInt()
        this.getData()
        var pos = 2
        for (d in 0 until ccv) {
            this.colorTable[d + 8] = java.awt.Color(if (data!![pos] < 0) 255 + data!![pos] else data!![pos], if (data!![pos + 1] < 0) 255 + data!![pos + 1] else data!![pos + 1], if (data!![pos + 2] < 0) 255 + data!![pos + 2] else data!![pos + 2])
            pos += 4
        }
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 157670739931392705L
    }

}