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

import java.util.Arrays

/**
 * SXPI 0x86
 *
 *
 * The SXPI record specifies the pivot fields and information about filtering on
 * the page axis of a PivotTable view. MUST exist if and only if the value of
 * the cDimPg field of the SxView record of the PivotTable view is greater than
 * zero.
 *
 *
 * rgsxpi (variable): An array of SXPI_Items that specifies the pivot fields and
 * information about filtering on the page axis of a PivotTable view. The number
 * of array elements MUST equal the value of the cDimPg field of the SxView
 * record of the PivotTable view.
 *
 *
 * The SXPI_Item structure specifies information about a pivot field and its
 * filtering on the page axis of a PivotTable view.
 *
 *
 * isxvd (2 bytes): A signed integer that specifies a pivot field index as
 * specified by Pivot Fields. The referenced pivot field is specified to be on
 * the page axis. MUST be greater than or equal to zero and less than the cDim
 * field of the SxView record of the PivotTable view.
 *
 *
 * isxvi (2 bytes): A signed integer that specifies the pivot item used for the
 * page axis filtering. MUST be a value from the following table: Value Meaning
 * 0x0000 to 0x7FFC This value specifies a pivot item index that specifies a
 * pivot item in the pivot field specified by isxvd. The referenced pivot item
 * specifies the page axis filtering for the pivot field. 0x7FFD This value
 * specifies all pivot items, see page axis for filtering that applies. For a
 * non-OLAP PivotTable view the value MUST be 0x7FFD or greater than or equal to
 * zero and less than the cItm field of the Sxvd record of the pivot field.
 * Otherwise the value MUST be 0x7FFD.
 *
 *
 * idObj (2 bytes): A signed integer that specifies the object identifier of the
 * Obj record with the page item drop-down arrow.
 */
class SxPI : XLSRecord(), XLSConstants {
    private var sxpis: Array<SXPI_Item>? = null

    override fun init() {
        super.init()
        val rgsxpi = this.getData()    // each item is 6 bytes
        if (rgsxpi != null) {
            if (rgsxpi.size % 6 != 0)
                Logger.logWarn("PivotTable: Irregular SxPI structure")
            sxpis = arrayOfNulls<SXPI_Item>(rgsxpi.size / 6)
            for (j in sxpis!!.indices) {
                sxpis[j] = SXPI_Item(rgsxpi, j * 6)
            }
            if (DEBUGLEVEL > 3) Logger.logInfo("SXPI - n: " + sxpis!!.size + ": " + Arrays.toString(data))
        } else if (DEBUGLEVEL > 3) Logger.logInfo("SXPI - NULL")
    }

    override fun toString(): String {
        return if (sxpis != null)
            "SXPI - n: " + sxpis!!.size + ": " + Arrays.toString(sxpis)
        else
            "SXPI - NULL"
    }

    /**
     * returns the pivot field index and item index for page axis field i
     *
     * @param i page axis field
     * @return
     */
    fun getPivotFieldItem(i: Int): IntArray {
        return if (i >= 0 && i < sxpis!!.size) {
            intArrayOf(sxpis!![i].isxvd.toInt(), sxpis!![i].idObj.toInt())
        } else intArrayOf(-1, -1)
    }

    /**
     * sets the pivot field index and item index for page axis field i
     *
     * @param i
     * @param fieldindex
     * @param itemindex
     */
    fun setPageFieldIndex(i: Int, fieldindex: Int, itemindex: Int) {
        if (i >= 0 && i < sxpis!!.size) {
            sxpis!![i].isxvd = fieldindex.toShort()
            sxpis!![i].isxvi = itemindex.toShort()
            var b = ByteTools.shortToLEBytes(fieldindex.toShort())
            this.getData()[i * 6] = b[0]
            this.getData()[i * 6 + 1] = b[1]
            b = ByteTools.shortToLEBytes(itemindex.toShort())
            this.getData()[i * 6 + 2] = b[0]
            this.getData()[i * 6 + 3] = b[1]
        }
    }

    /**
     * add a pivot field to the page axis
     *
     * @param fieldIndex
     * @param itemIndex
     */
    fun addPageField(fieldIndex: Int, itemIndex: Int) {
        getData()
        data = ByteTools.append(ByteArray(6), data)
        val tmp = arrayOfNulls<SXPI_Item>(sxpis!!.size + 1)
        System.arraycopy(sxpis!!, 0, tmp, 0, sxpis!!.size)
        sxpis = tmp
        sxpis[sxpis!!.size - 1] = SXPI_Item()

        setPageFieldIndex(sxpis!!.size - 1, fieldIndex, itemIndex)

    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 2639291289806138985L

        // no data, initially???
        val prototype: XLSRecord?
            get() {
                val sp = SxPI()
                sp.opcode = XLSConstants.SXPI
                sp.setData(byteArrayOf())
                sp.init()
                return sp
            }
    }

}

/**
 * helper class defines SxPI structure
 */
internal class SXPI_Item {
    var isxvd: Short = 0    // pivot field index
    var isxvi: Short = 0    // specifies the pivot item used for the page axis filtering; if 0x7FFD This valuespecifies all pivot items, otherwise it's item index of pivot field (isxvd)
    var idObj: Short = 0

    constructor() {}

    constructor(rgsxpi: ByteArray, idx: Int) {
        var idx = idx
        isxvd = ByteTools.readShort(rgsxpi[idx++].toInt(), rgsxpi[idx++].toInt())
        isxvi = ByteTools.readShort(rgsxpi[idx++].toInt(), rgsxpi[idx++].toInt())
        idObj = ByteTools.readShort(rgsxpi[idx++].toInt(), rgsxpi[idx++].toInt())
    }
}