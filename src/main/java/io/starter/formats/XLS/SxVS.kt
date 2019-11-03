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
 * SXVS  0xE3
 * The SXVS record specifies the type of source data used for a PivotCache.
 * This record is followed by a sequence of records that specify additional information about the source data.
 *
 *
 * sxvs (2 bytes): An unsigned integer that specifies the type of source data used for the PivotCache. The types of records that follow this record are dictated by the value of this field. MUST be a value from the following table:
 *
 *
 * Name		Value		Meaning
 *
 *
 * SHEET		0x0001		Specifies that the source data is a range. This record MUST be followed by a DConRef record that specifies a simple range, or a DConName record that specifies a named range, or a DConBin record that specifies a built-in named range.
 * EXTERNAL	0x0002		Specifies that external source data is used. This record MUST be followed by a sequence of records beginning with a DbQuery record that specifies connection and query information that is used to retrieve external data.
 * CONSOLIDATION	0x0004	Specifies that multiple consolidation ranges are used as the source data. This record MUST be followed by a sequence of records beginning with an SXTbl record that specifies information about the multiple consolidation ranges.
 * SCENARIO	0x0010		The source data is populated from a temporary internal structure. In this case there is no additional source data information because the raw data does not exist as a permanent structure and the logic to produce it is application-dependent.
 */
class SxVS : XLSRecord(), XLSConstants {
    /**
     * serialVersionUID
     */
    var sourceType: Short = -1
        private set

    override fun init() {
        super.init()
        sourceType = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        if (DEBUGLEVEL > 3) Logger.logInfo("SXVS - sourceType:$sourceType")
    }

    fun setSourceType(st: Int) {
        sourceType = st.toShort()
        val b = ByteTools.shortToLEBytes(sourceType)
        this.getData()[0] = b[0]
        this.getData()[1] = b[1]
    }

    companion object {
        val TYPE_SHEET: Short = 0x1
        val TYPE_EXTERNAL: Short = 0x0002
        val TYPE_CONSOLIDATION: Short = 0x0004
        val TYPE_SCENARIO: Short = 0x0010

        private val serialVersionUID = 2639291289806138985L

        val prototype: XLSRecord?
            get() {
                val sv = SxVS()
                sv.opcode = XLSConstants.SXVS
                sv.setData(byteArrayOf(1, 0))
                sv.init()
                return sv
            }
    }
}
