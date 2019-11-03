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
 * **Fontx: Font Index (0x1026)**
 * Child of a TEXT record and defines a text font by indexing the appropriate font in
 * the font table.
 *
 *
 * 4		iFont		2		Index number into the font table
 * If this field is less than or equal to the number of Font records
 * in the workbook, this field is a one-based index to a Font record in
 * the workbook. Otherwise, this field is a one-based index into the
 * collection of Font records in the chart sheet substream, where the index is equal to
 * iFont minus n and n is the number of Font records in the workbook.
 */
class Fontx : GenericChartObject(), ChartObject {
    // 20070806 KSC: Add init/update to control FontX opts
    private var ifnt: Short = 0

    private val PROTOTYPE_BYTES = byteArrayOf(5, 0)

    override fun init() {
        super.init()
        ifnt = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
    }

    /**
     * returns the index into the wb font table referenced by this Fontx record for the chart
     */
    /*
	If this field is less than or equal to the number of Font records
	in the workbook, this field is a one-based index to a Font record in
	the workbook. Otherwise, this field is a one-based index into the
	collection of Font records in the chart sheet substream, where the index is equal to
	iFont minus n and n is the number of Font records in the workbook.
	*/
    fun getIfnt(): Int {
        var n = 0
        try {
            n = this.workBook!!.numFonts
        } catch (e: Exception) {
        }

        return if (ifnt <= n)
            ifnt.toInt()
        else
            ifnt - n + 1
    }

    fun setIfnt(id: Int) {
        this.ifnt = id.toShort()
        val b = ByteTools.shortToLEBytes(this.ifnt)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 4255798925225768809L

        val prototype: XLSRecord?
            get() {
                val f = Fontx()
                f.opcode = XLSConstants.FONTX
                f.data = f.PROTOTYPE_BYTES
                f.init()
                return f
            }
    }
}
