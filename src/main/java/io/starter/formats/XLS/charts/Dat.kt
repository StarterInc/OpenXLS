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
 * ** Dat:  Data Table Options  (0x1063)**
 *
 *
 * Offset	Name	Size	Contents
 * 4		grbit	2		Option flags (see following table)
 *
 *
 *
 *
 * The grbit field contains the following flags.
 *
 *
 * Offset	Bits	Mask	Name			Contents
 * 0		0		01h		fHasBordHorz	1 = data table has horizontal borders
 * 1		02h		fHasBordVert	1 = data table has vertical borders
 * 2		04h		fhasBordOutline	1 = data table has a border
 * 3		08h		fShowSeriesKey	1 = data table shows series keys
 * 7-4		F0h		reserved		Reserved; must be zero
 * 1		7-0		FFh		reserved		Reserved; must be zero
 */
class Dat : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    internal var fHasBordHorz: Boolean = false
    internal var fHasBordVert: Boolean = false
    internal var fHasBordOutline: Boolean = false
    internal var fShowSeriesKey: Boolean = false

    private val PROTOTYPE_BYTES = byteArrayOf(15, 0)

    override fun init() {
        super.init()
        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        fHasBordHorz = grbit and 0x1 == 0x1
        fHasBordVert = grbit and 0x2 == 0x2
        fHasBordOutline = grbit and 0x4 == 0x4
        fShowSeriesKey = grbit and 0x8 == 0x8

    }

    private fun updateRecord() {
        val b = ByteTools.shortToLEBytes(grbit)
        this.data = b
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 1138056714558134785L

        /**
         * creates a new Dat record; if bCreateDataTable is true,
         * will also add the associated records required to create
         * a Data Table
         *
         * @param bCreateDataTable
         * @return
         */
        fun getPrototype(bCreateDataTable: Boolean): XLSRecord {
            val d = Dat()
            d.opcode = XLSConstants.DAT
            d.data = d.PROTOTYPE_BYTES
            d.init()
            if (bCreateDataTable) {
                val l = Legend.prototype as Legend?
                l!!.setIsDataTable(true)
                //l.setwType(Legend.NOT_DOCKED);
                d.chartArr.add(l)
                // add pos record
                var p = Pos.getPrototype(Pos.TYPE_DATATABLE) as Pos
                //            p.setWorkBook(book);
                l.addChartRecord(p)
                // TextDisp
                val td = TextDisp.prototype as TextDisp?
                l.addChartRecord(td)
                // TextDisp sub-recs
                p = Pos.getPrototype(Pos.TYPE_TEXTDISP) as Pos
                td!!.addChartRecord(p)
                val f = Fontx.prototype as Fontx?
                //	        f.setWorkBook(book);
                // EVENTUALLY!
                f!!.ifnt = 6
                td.addChartRecord(f)
                val ai = Ai.getPrototype(Ai.AI_TYPE_NULL_LEGEND) as Ai
                //            ai.setWorkBook(book);
                td.addChartRecord(ai)
            }
            return d
        }
    }
}
