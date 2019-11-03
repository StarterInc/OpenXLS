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

import java.io.Serializable

internal abstract class EscherRecord(var fbt: Int, var inst: Int, var version: Int) : Serializable {
    var header: ByteArray? = null
    var data: ByteArray? = null
    var length: Int = 0
    var isDirty = false

    private//TODO: Reverse the process of header decoding here
    val headerBytes: ByteArray
        get() {
            val headerBytes = ByteArray(4)

            headerBytes[0] = (0xF and version or (0xF0 and (inst shl 4))).toByte()
            headerBytes[1] = (0x00000FF0 and inst shr 4).toByte()
            headerBytes[2] = (0x000000FF and fbt).toByte()
            headerBytes[3] = (0x0000FF00 and fbt shr 8).toByte()


            val version2 = 0x0F and headerBytes[0]
            val inst2 = 0xFF and headerBytes[1] shr 4 or (0xF0 and headerBytes[0] shr 4)

            val lenBytes = ByteTools.cLongToLEBytes(length)

            val retData = ByteArray(8)
            System.arraycopy(headerBytes, 0, retData, 0, 4)
            System.arraycopy(lenBytes, 0, retData, headerBytes.size, 4)

            return retData
        }

    /**
     * no param constructor for Serializable
     */
    fun EscherRecord() {}

    protected abstract fun getData(): ByteArray

    fun toByteArray(): ByteArray {
        val dataBytes = getData()  //Have it in this sequence as some records adjust their header byte length from getData
        val headerBytes = headerBytes

        val retData = ByteArray(headerBytes.size + dataBytes.size)
        System.arraycopy(headerBytes, 0, retData, 0, headerBytes.size)
        if (dataBytes.size > 0)
            System.arraycopy(dataBytes, 0, retData, headerBytes.size, dataBytes.size)
        return retData

    }

    companion object {

        /**
         * serialVersionUID
         */
        private const val serialVersionUID = -7987132917889379656L
    }
}
