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


//0xF00A
class MsofbtSp(fbt: Int, inst: Int, version: Int) : EscherRecord(fbt, inst, version) {
    internal var id: Int = 0
    internal var grfPersistence: Int = 0

    public override fun getData(): ByteArray {
        val idBytes: ByteArray
        val flagBytes: ByteArray
        idBytes = ByteTools.cLongToLEBytes(id)
        flagBytes = ByteTools.cLongToLEBytes(grfPersistence)
        val retData = ByteArray(8)
        System.arraycopy(idBytes, 0, retData, 0, 4)
        System.arraycopy(flagBytes, 0, retData, 4, 4)

        this.length = 8
        return retData

    }

    fun setId(value: Int) {

        id = value
    }

    fun setGrfPersistence(value: Int) {
        grfPersistence = value
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -5355585244930369889L
    }
}
