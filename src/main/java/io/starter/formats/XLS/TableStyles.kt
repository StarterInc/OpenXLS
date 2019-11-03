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

import java.io.UnsupportedEncodingException

/**
 * TableStyles 0x88E
 *
 *
 * The TableStyles record specifies the default table and PivotTabletable styles and
 * specifies the beginning of a collection of TableStyle records as defined by the Globals substream.
 * The collection of TableStyle records specifies user-defined table styles.
 *
 *
 * frtHeader (12 bytes): An FrtHeader structure. The frtHeader.rt field MUST be 0x088E.
 *
 *
 * cts (4 bytes): An unsigned integer that specifies the total number of table styles in this document. This is the sum of the standard built-in table styles and all of the custom table styles. This value MUST be greater than or equal to 144 (the number of built-in table styles).
 *
 *
 * cchDefTableStyle (2 bytes): An unsigned integer that specifies the count of characters in the rgchDefTableStyle field. This value MUST be less than or equal to 255.
 *
 *
 * cchDefPivotStyle (2 bytes): An unsigned integer that specifies the count of characters in the rgchDefPivotStyle field. This value MUST be less than or equal to 255.
 *
 *
 * rgchDefTableStyle (variable): An array of Unicode characters whose length is specified by cchDefTableStyle that specifies the name of the default table style.
 *
 *
 * rgchDefPivotStyle (variable): An array of Unicode characters whose length is specified by cchDefPivotStyle that specifies the name of the default PivotTable style.
 */
class TableStyles : XLSRecord(), XLSConstants {
    /**
     * serialVersionUID
     */
    internal var cts: Short = 0
    internal var cchDefTableStyle: Short = 0
    internal var cchDefPivotStyle: Short = 0
    internal var rgchDefTableStyle: String? = null
    internal var rgchDefPivotStyle: String? = null

    override fun init() {
        super.init()
        // -114, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,  // old frtHeader
        //-112, 0, 0, 0, 	== 144
        // An unsigned integer that specifies the total number of table styles in this document. This is the sum of the standard built-in table styles and all of the custom table styles. This value MUST be greater than or equal to 144 (the number of built-in table styles).
        cts = ByteTools.readInt(this.getByteAt(12), this.getByteAt(13), this.getByteAt(14), this.getByteAt(14)).toShort()
        cchDefTableStyle = ByteTools.readShort(this.getByteAt(16).toInt(), this.getByteAt(17).toInt())
        cchDefPivotStyle = ByteTools.readShort(this.getByteAt(18).toInt(), this.getByteAt(19).toInt())
        var pos = 20
        if (cchDefTableStyle > 0) {
            val tmp = this.getBytesAt(pos, cchDefTableStyle * 2)
            try {
                rgchDefTableStyle = String(tmp!!, XLSConstants.UNICODEENCODING)
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("encoding Table Style name in TableStyles: $e")
            }

            pos += cchDefTableStyle * 2
        }
        if (cchDefPivotStyle > 0) {
            val tmp = this.getBytesAt(pos, cchDefPivotStyle * 2)
            try {
                rgchDefPivotStyle = String(tmp!!, XLSConstants.UNICODEENCODING)
            } catch (e: UnsupportedEncodingException) {
                Logger.logInfo("encoding Pivot Style name in TableStyles: $e")
            }

        }
    }

    companion object {
        private val serialVersionUID = 2639291289806138985L

        /* required id *//* cts *//* cch *//* cch */ val prototype: XLSRecord?
            get() {
                val tx = TableStyles()
                tx.opcode = XLSConstants.TABLESTYLES
                tx.setData(byteArrayOf(-114, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -112, 0, 0, 0, 17, 0, 17, 0, 84, 0, 97, 0, 98, 0, 108, 0, 101, 0, 83, 0, 116, 0, 121, 0, 108, 0, 101, 0, 77, 0, 101, 0, 100, 0, 105, 0, 117, 0, 109, 0, 57, 0, 80, 0, 105, 0, 118, 0, 111, 0, 116, 0, 83, 0, 116, 0, 121, 0, 108, 0, 101, 0, 76, 0, 105, 0, 103, 0, 104, 0, 116, 0, 49, 0, 54, 0))
                tx.init()
                return tx
            }
    }
}
