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

/**
 * **SXVIEWLINK: Chart PivotTable Name (858h)**
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record
 * for Charts. This record stores the name of the source PivotTable
 * when this chart is a PivotChart. New for Excel 9, PivotCharts are
 * charts based on PivotTables.
 *
 *
 * Record Data
 * Offset	Field Name	Size	Contents
 * 4		rt			2		Record type; this matches the BIFF rt in the first two bytes of the record; =0858h
 * 6		grbitFrt	2		FRT flags; must be zero
 * 8		brst		var		String containing name of PivotTable
 */
class SxViewLink : GenericChartObject(), ChartObject {

    // TODO: Prototype Bytes
    private val PROTOTYPE_BYTES = byteArrayOf()

    override fun init() {
        super.init()
    }

    private fun updateRecord() {}

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -5291795207491688189L

        val prototype: XLSRecord?
            get() {
                val sx = SxViewLink()
                sx.opcode = XLSConstants.SXVIEWLINK
                sx.data = sx.PROTOTYPE_BYTES
                sx.init()
                return sx
            }
    }

}
