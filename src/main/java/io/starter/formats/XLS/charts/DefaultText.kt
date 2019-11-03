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

package io.starter.formats.XLS.charts

import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools

/**
 * **DefaultText: Default Data Label Text Properties(0x1024)**
 *
 *
 * id (2 bytes): An unsigned integer that specifies the text elements that are formatted using the position and appearance information specified by the Text record immediately following this record. MUST be a value from the following table.
 * If this record is in a sequence of records that conforms to the CRT rule as specified by the Chart Sheet Substream ABNF, then this field MUST be 0x0000 or 0x0001. If this record is not in a sequence of records that conforms to the CRT rule as specified by the Chart Sheet Substream ABNF, then this field MUST be 0x0002 or 0x0003.
 * Value		Meaning
 * 0x0000		Format all Text records in the chart group where fShowPercent is equal to 0 or fShowValue is equal to 0.
 * 0x0001		Format all Text records in the chart group where fShowPercent is equal to 1 or fShowValue is equal to 1.
 * 0x0002		Format all Text records in the chart where the value of fScaled of the associated FontInfo structure is equal to 0.
 * 0x0003		Format all Text records in the chart where the value of fScaled of the associated FontInfo structure is equal to 1. *
 */
class DefaultText : GenericChartObject(), ChartObject {

    private var grbit: Short = 0

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0)

    // try to interpret!!!
    val type: Int
        get() = grbit.toInt()

    override fun init() {
        super.init()
        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
    }

    fun setType(type: Short) {
        grbit = type
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 6845131434077760152L

        // 20070716 KSC: Need to create new records
        val prototype: XLSRecord?
            get() {
                val t = DefaultText()
                t.opcode = XLSConstants.DEFAULTTEXT
                t.data = t.PROTOTYPE_BYTES
                t.init()
                return t
            }

        val TYPE_SHOWLABELS = 0
        val TYPE_VALUELABELS = 1
        val TYPE_ALLTEXT = 2
        val TYPE_UNKNOWN = 3
    }

}
