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
 * **CATLAB: Category Labels (856h)**
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record for Charts.
 *
 *
 * Record Data
 * Offset		Field Name	Size	Contents
 * 10			wOffset		2		Distance between category label levels
 * 12			at			2		Category axis label alignment
 * 14			grbit		2		Option flags for category axis labels (see description below)
 * 16			(unused)	2		Reserved; must be zero
 *
 *
 * The grbit field contains the following category axis label option flags:
 * Bits	Mask	Flag Name			Contents
 * 0		0001h	fAutoCatLabelReal	=1 if the category label skip is automatic =0 otherwise
 * 15-1	FFFEh	(unused)			Reserved; must be zero
 *
 *
 *
 *
 * wOffset (2 bytes): An unsigned integer that specifies the distance between the axis and axis label.
 * It contains the offset as a percentage of the default distance. The default distance is equal to 1/3 the height of the font calculated in pixels.
 * MUST be a value greater than or equal to 0 (0%) and less than or equal to 1000 (1000%).
 * at (2 bytes): An unsigned integer that specifies the alignment of the axis label. MUST be a value from the following table:
 * Value			Alignment
 * 0x0001			Top-aligned if the trot field of the Text record of the axis is not equal to 0. Left-aligned if the iReadingOrder field of the Text record of the axis specifies left-to-rightreading order; otherwise, right-aligned.
 * 0x0002		    Center-alignment
 * 0x0003			Bottom-aligned if the trot field of the Text record of the axis is not equal to 0. Right-aligned if the iReadingOrder field of the Text record of the axis specifies left-to-right reading order; otherwise, left-aligned.
 *
 *
 * A - cAutoCatLabelReal (1 bit): A bit that specifies whether the number of categories (3) between axis labels is set to the default value. MUST be a value from the following table:
 * Value	Description
 * 0	    The value is set to catLabel field as specified by CatSerRange record.
 * 1	    The value is set to the default value. The number of category (3) labels is automatically calculated by the application based on the data in the chart.
 * unused (15 bits): Undefined, and MUST be ignored.
 * reserved (2 bytes): MUST be zero, and MUST be ignored.
 */
class CatLab : GenericChartObject(), ChartObject {
    /**
     * serialVersionUID
     */
    internal var wOffset: Short = 0
    internal var at: Short = 0

    // TODO:  need prototype bytes
    private val PROTOTYPE_BYTES = byteArrayOf(86.toByte(), 8, 0, 0, 100.toByte(), 0, 2, 0, 86, 66, 0, 0)

    override fun init() {
        super.init()
        wOffset = ByteTools.readShort(this.data!![4].toInt(), this.data!![5].toInt())
        at = ByteTools.readShort(this.data!![6].toInt(), this.data!![7].toInt())
        // cAutoCatLabelReal= (record[8] & 0x1==0x1
    }

    fun setOption(op: String, `val`: String) {
        if (op == "lblAlign") {            // ctr, l, r
            if (`val` == "ctr")
                at = 2
            else if (`val` == "l")
                at = 1
            else
                at = 3
        } else if (op == "lblOffset") {    // 0-100
            wOffset = Integer.parseInt(`val`).toShort()
        }
        updateRecord()
    }

    fun getOption(op: String): String? {
        if (op == "lblAlign") {            // ctr, l, r
            return if (at.toInt() == 2)
                "ctr"
            else if (at.toInt() == 1)
                "l"
            else
                "r"
        } else if (op == "lblOffset")
        // 0-100
            return Integer.toString(wOffset.toInt())
        return null
    }

    private fun updateRecord() {
        var b = ByteArray(2)
        b = ByteTools.shortToLEBytes(wOffset)
        this.data[4] = b[0]
        this.data[5] = b[1]
        b = ByteTools.shortToLEBytes(at)
        this.data[6] = b[0]
        this.data[7] = b[1]
    }

    companion object {
        private val serialVersionUID = 3042712098138741496L

        val prototype: XLSRecord?
            get() {
                val cl = CatLab()
                cl.opcode = XLSConstants.CATLAB
                cl.data = cl.PROTOTYPE_BYTES
                cl.init()
                return cl
            }
    }

}
