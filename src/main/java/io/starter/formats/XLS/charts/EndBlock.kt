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
 * **ENDBLOCK: Chart Future Record Type End Block (853h)**
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record for
 * Charts that indicates end of an object's scope for Pre-Excel 9 objects.
 * Paired with STARTBLOCK.
 *
 *
 * Record Data
 * Offset		Field Name		Size		Contents
 * 4			rt				2			Record type; this matches the BIFF rt in the first two bytes of the record; =0853h
 * 6			grbitFrt		2			FRT flags; must be zero
 * 8			iObjectKind		2			Sanity check for object scope being ended
 * 10			(unused)		6			Reserved; must be zero
 */
class EndBlock : GenericChartObject(), ChartObject {
    internal var iObjectKind: Short = 0

    private val PROTOTYPE_BYTES = byteArrayOf(83, 8, 0, 0, 5, 0, 0, 0, 0, 0, 0, 0)

    override fun init() {
        super.init()
        iObjectKind = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
    }    // iObjectKind= 5 or 0 or 13

    fun setObjectKind(i: Int) {
        iObjectKind = i.toShort()
        updateRecord()
    }

    private fun updateRecord() {
        val b = ByteTools.shortToLEBytes(iObjectKind)
        this.data[4] = b[0]
        this.data[5] = b[1]
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 1132544432743236942L

        val prototype: XLSRecord?
            get() {
                val eb = EndBlock()
                eb.opcode = XLSConstants.ENDBLOCK
                eb.data = eb.PROTOTYPE_BYTES
                eb.init()
                return eb
            }
    }

}
