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
package io.starter.formats.escher

import io.starter.toolkit.ByteTools


//0xF008
class MsofbtDg(fbt: Int, inst: Int, version: Int) : EscherRecord(fbt, inst, version) {
    internal var csp: Int = 0
    internal var lastSPID: Int = 0

    override fun getData(): ByteArray {
        val cspBytes: ByteArray
        val spIdBytes: ByteArray

        cspBytes = ByteTools.cLongToLEBytes(csp)        // Number of shapes
        spIdBytes = ByteTools.cLongToLEBytes(lastSPID)    // last SPID

        val retBytes = ByteArray(8)
        System.arraycopy(cspBytes, 0, retBytes, 0, 4)
        System.arraycopy(spIdBytes, 0, retBytes, 4, 4)

        this.length = retBytes.size
        return retBytes
    }

    fun setNumShapes(value: Int) {
        csp = value
    }

    fun setLastSPID(value: Int) {

        lastSPID = value
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 5218802290529676567L
    }
}
