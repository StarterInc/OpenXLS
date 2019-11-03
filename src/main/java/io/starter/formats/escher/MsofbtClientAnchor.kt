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

import java.io.ByteArrayOutputStream

//0xf010
class MsofbtClientAnchor(fbt: Int, inst: Int, version: Int) : EscherRecord(fbt, inst, version) {
    // 20070914 KSC: this record doesn't store sheet index; 1st two bytes are a flag (seems always to be 2)
    internal var flag: Short = 2
    internal var leftColumnIndex: Short = 0
    internal var xOffsetL: Short = 0
    internal var topRowIndex: Short = 0
    internal var yOffsetT: Short = 0
    internal var rightColIndex: Short = 0
    internal var xOffsetR: Short = 0
    internal var bottomRowIndex: Short = 0
    internal var yOffsetB: Short = 0

    public override fun getData(): ByteArray {
        val bos = ByteArrayOutputStream()
        try {
            bos.write(ByteTools.shortToLEBytes(flag))
            bos.write(ByteTools.shortToLEBytes(leftColumnIndex))
            bos.write(ByteTools.shortToLEBytes(xOffsetL))
            bos.write(ByteTools.shortToLEBytes(topRowIndex))
            bos.write(ByteTools.shortToLEBytes(yOffsetT))
            bos.write(ByteTools.shortToLEBytes(rightColIndex))
            bos.write(ByteTools.shortToLEBytes(xOffsetR))
            bos.write(ByteTools.shortToLEBytes(bottomRowIndex))
            bos.write(ByteTools.shortToLEBytes(yOffsetB))
        } catch (e: Exception) {

        }

        this.length = bos.toByteArray().size
        return bos.toByteArray()

    }

    fun setBounds(bounds: ShortArray?) {
        if (bounds == null) return
        leftColumnIndex = bounds[0]
        xOffsetL = bounds[1]
        topRowIndex = bounds[2]
        yOffsetT = bounds[3]
        rightColIndex = bounds[4]
        xOffsetR = bounds[5]
        bottomRowIndex = bounds[6]
        yOffsetB = bounds[7]

    }

    fun setFlag(flag: Short) {
        this.flag = flag
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 7946989934447191433L
    }

}
