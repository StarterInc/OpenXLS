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

class MsofbtSplitMenuColors(fbt: Int, inst: Int, version: Int) : EscherRecord(fbt, inst, version) {
    //These values are from experimental records.
    var fillColor = 0x800000D
    var lineColor = 0x800000C
    var shadowColor = 0x8000017
    var 3dColor = 0x100000f7

    // @Override
    override fun getData(): ByteArray {
        val fillColorBytes = ByteTools.cLongToLEBytes(fillColor)
        val lineColorBytes = ByteTools.cLongToLEBytes(lineColor)
        val shadowColorBytes = ByteTools.cLongToLEBytes(shadowColor)
        val _3dColorBytes = ByteTools.cLongToLEBytes(3 dColor)

        val totalBytes = ByteArray(16)

        System.arraycopy(fillColorBytes, 0, totalBytes, 0, 4)
        System.arraycopy(lineColorBytes, 0, totalBytes, 4, 4)
        System.arraycopy(shadowColorBytes, 0, totalBytes, 8, 4)
        System.arraycopy(_3dColorBytes, 0, totalBytes, 12, 4)

        this.length = 16
        this.inst = 4

        return totalBytes
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 5888748984363726576L
    }

}
