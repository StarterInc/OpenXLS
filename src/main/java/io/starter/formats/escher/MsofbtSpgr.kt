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

//0xf009

/*
 * This record is present only in group shapes (not shapes in groups, shapes that are groups).
 * The group shape record defines the coordinate system of the shape, which the anchors of the
 * child shape are expressed in. All other information is stored in the shape records that follow.
 */
class MsofbtSpgr(fbt: Int, inst: Int, version: Int) : EscherRecord(fbt, inst, version) {
    internal var left = 0
    internal var top = 0
    internal var right = 0
    internal var bottom = 0

    override fun getData(): ByteArray {
        val leftBytes: ByteArray
        val topBytes: ByteArray
        val rightBytes: ByteArray
        val bottomBytes: ByteArray
        leftBytes = ByteTools.cLongToLEBytes(left)
        topBytes = ByteTools.cLongToLEBytes(top)
        rightBytes = ByteTools.cLongToLEBytes(right)
        bottomBytes = ByteTools.cLongToLEBytes(bottom)

        val retBytes = ByteArray(16)
        System.arraycopy(leftBytes, 0, retBytes, 0, 4)
        System.arraycopy(topBytes, 0, retBytes, 4, 4)
        System.arraycopy(rightBytes, 0, retBytes, 8, 4)
        System.arraycopy(bottomBytes, 0, retBytes, 12, 4)

        this.length = retBytes.size

        return retBytes
    }

    fun setRect(left: Int, top: Int, right: Int, bottom: Int) {
        this.left = left
        this.right = right
        this.bottom = bottom
        this.top = top

    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 5591214948365806058L
    }

}
